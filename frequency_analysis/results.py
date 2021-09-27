'''Additional module to frequency.py for excel output.'''

import os
import re
import sqlite3
from string import ascii_lowercase
import xlsxwriter


class ExcelWriter:
    '''Convert generated .db data to excel view.

    All mandatory functions are called all at once by treat().
    Additional functions – sheet_en_symb_bigrams(), sheet_ru_symb_bigrams() and sheet_yo_words()
        are called individually.
    '''

    def __init__(self, workbook, cursor):
        self.workbook = workbook
        self.cursor = cursor
        self.f_bold = self.workbook.add_format({'bold': True, 'align': 'center'})
        self.f_percent = self.workbook.add_format({'num_format': '0.00%', 'align': 'center'})
        self.f_int = self.workbook.add_format({'num_format': '#,##0', 'align': 'center'})
        self.f_float = self.workbook.add_format({'num_format': '#,##0.00', 'align': 'center'})
        self.f_red_bg = self.workbook.add_format({'bg_color': '#FFC7CE', 'align': 'center'})

        self.pos_list = [
            self.cursor.execute(f'SELECT AVG(position) FROM {x};').fetchone()[0]
            for x in ('symbols', 'symbol_bigrams', 'words', 'word_bigrams')
        ]
        self.sum_list = [
            self.cursor.execute(f'SELECT SUM(quantity) FROM {x};').fetchone()[0]
            for x in ('symbols', 'symbol_bigrams', 'words', 'word_bigrams')
        ]

    def __add_main_style(self, sheet, f_width=5, a_width=12, two_column=False):
        '''Add main row/column formating (width, bold, centred, freeze).'''
        sheet.ignore_errors({'number_stored_as_text': 'A:ZZ'})
        sheet.freeze_panes(1, 1 + two_column)
        sheet.set_row(0, None, self.f_bold)
        sheet.set_column('A:A', f_width, self.f_bold)
        if two_column:
            sheet.set_column('C:AZ', a_width)
            sheet.set_column('B:B', f_width, self.f_bold)
        else:
            sheet.set_column('B:AZ', a_width)

    def treat(self, limits=(0, 0, 0, 0), min_quantity=(1, 1, 1, 1, 1)):
        '''Create main sheets all at once.

        Input:
            limits  – tuple – max number of elements to be added to the sheet (0 – unlimited))
                    – (symbols, symbol bigrams top, words top, word bigrams top)
                    – default values – [0, 0, 0, 0] (ommited = default);
            min_quantity – tuple – min number of entries for each element to take it into account)
                    – (/same as on 'limits'/, +symbol bigrams table)
                    – default values – (1, 1, 1, 1, 1) (ommited = default)
                    – (symbs, symb bigrs top, words top, word bigrs top, !symbol bigrams table!).
        '''
        if len(limits) < 4:
            limits = list(limits) + [0] * (4 - len(limits))
        if len(min_quantity) < 5:
            min_quantity = list(min_quantity) + [1] * (5 - len(min_quantity))
        print('Start of writing to .xlsx')
        self.sheet_stats()
        print('... stats sheet was written')
        self.sheet_symbols(limits[0], min_quantity[0])
        print('... symbols sheet was written')
        self.sheet_symbol_bigrams_top(limits[1], min_quantity[1])
        print('... top symbol bigrams sheet was written')
        self.sheet_all_symb_bigrams(min_quantity[4])
        print('... symbol bigrams table sheet was written')
        self.sheet_top_words(limits[2], min_quantity[2])
        print('... top words sheet was written')
        self.sheet_word_bigrams_top(limits[3], min_quantity[3])
        print('... top word bigrams sheet was written')
        print('End of writing main sheets.')
        print(
            'You can call additional functions to create more sheets '
            '(e.g. "sheet_en_symb_bigrams()", "sheet_ru_symb_bigrams()", '
            '"sheet_yo_words([limit], [min_quantity])") or "sheet_custom_symb(symbols_string)".'
        )

    def sheet_stats(self):
        '''Create main statistic of analysis. Is called from main "treat()".'''
        count_list = [
            self.cursor.execute(f'SELECT COUNT(*) FROM {x};').fetchone()[0]
            for x in ('symbols', 'symbol_bigrams', 'words', 'word_bigrams')
        ]
        avg_pos_list = [
            self.cursor.execute(
                f'SELECT SUM(quantity*position)/SUM(quantity) FROM {x} \
                        {"WHERE ord != 32" if x == "symbols" else ""};'
            ).fetchone()[0]
            for x in ('symbols', 'symbol_bigrams', 'words', 'word_bigrams')
        ]
        stats = self.workbook.add_worksheet('Stats')
        self.__add_main_style(stats, 15, 15)
        stats.write(0, 1, 'Total')
        stats.write(0, 2, 'Quantity')
        stats.write(0, 3, 'Avg. position')
        stats.write_column(1, 0, ('Symbols', 'Symbol bigrams', 'Words', 'Word bigrams'))
        stats.write_column(1, 1, count_list, self.f_int)
        stats.write_column(1, 2, self.sum_list, self.f_int)
        stats.write_column(1, 3, avg_pos_list, self.f_float)

    def sheet_symbols(self, limit=0, min_quantity=1):
        '''Create top-list of all analyzed symbols by quantity. Is called from main "treat()".'''
        symbols = self.workbook.add_worksheet('Symbols')
        symbols.set_tab_color('green')
        self.__add_main_style(symbols)
        symbols.write_row(0, 0, ('Symb', 'Quantity', '% from all', 'As first', 'As last'))
        if self.pos_list[0] != 1:
            symbols.write(0, 5, 'Avg. position')

        self.cursor.execute(
            f'''
            SELECT *
            FROM symbols
            WHERE quantity >= {min_quantity}
            ORDER BY quantity DESC, ord ASC
            {f'LIMIT {limit}' if limit else ''};
            '''
        )
        for row, symb in enumerate(self.cursor.fetchall(), 1):
            symbols.write_string(row, 0, chr(symb[0]))
            symbols.write_number(row, 1, symb[1], self.f_int)
            symbols.write_number(row, 2, symb[1] / self.sum_list[0], self.f_percent)
            if symb[0] != 32:
                symbols.write_number(row, 3, symb[2], self.f_int)
                symbols.write_number(row, 4, symb[3], self.f_int)
                if self.pos_list[0] != 1:
                    symbols.write_number(row, 5, symb[4], self.f_float)

        chart = self.workbook.add_chart({'type': 'pie'})
        chart.add_series(
            {
                'name': 'Letter frequency',
                'categories': f'=Symbols!$A3:$A{limit+3 if limit else ""}',
                'values': f'=Symbols!$C3:$C{limit+3 if limit else ""}',
            }
        )
        symbols.insert_chart('H2', chart)

    def sheet_symbol_bigrams_top(self, limit=0, min_quantity=1):
        '''Create top-list of symbol bigrams by quantity. Is called from main "treat()".'''
        symbol_bigrams_top = self.workbook.add_worksheet('Symbol bigrams top')
        symbol_bigrams_top.set_tab_color('green')
        self.__add_main_style(symbol_bigrams_top, two_column=True)
        symbol_bigrams_top.write_row(0, 0, ('1st', '2nd', 'Quantity', '% from all'))
        if self.pos_list[1]:
            symbol_bigrams_top.write(0, 4, 'Avg. position')

        self.cursor.execute(
            f'''
            SELECT *
            FROM symbol_bigrams
            WHERE quantity >= {min_quantity}
            ORDER BY quantity DESC, first_symb_ord ASC, second_symb_ord ASC
            {f'LIMIT {limit}' if limit else ''};
            '''
        )
        for row, bigr in enumerate(self.cursor.fetchall(), 1):
            symbol_bigrams_top.write_string(row, 0, chr(bigr[0]))
            symbol_bigrams_top.write_string(row, 1, chr(bigr[1]))
            symbol_bigrams_top.write_number(row, 2, bigr[2], self.f_int)
            symbol_bigrams_top.write_number(row, 3, bigr[2] / self.sum_list[1], self.f_percent)
            if self.pos_list[1]:
                symbol_bigrams_top.write_number(row, 4, bigr[3], self.f_float)

    def sheet_all_symb_bigrams(self, min_quantity=1):
        '''Create 2D bigrams table for all analyzed symbols. Is called from main "treat()".'''
        all_symb_bigrams = self.workbook.add_worksheet('All symb bigrams')
        all_symb_bigrams.set_tab_color('red')
        self.__add_main_style(all_symb_bigrams, 2.14, 9.43)
        self.cursor.execute(
            f'''
            SELECT *
            FROM symbol_bigrams
            WHERE quantity >= {min_quantity}
            ORDER BY first_symb_ord ASC, second_symb_ord ASC;
            '''
        )
        order: dict = {}
        for pair in self.cursor.fetchall():
            for elem_num in [0, 1]:
                if pair[elem_num] not in order:
                    order[pair[elem_num]] = len(order) + 1
                    all_symb_bigrams.write_string(0, order[pair[elem_num]], chr(pair[elem_num]))
                    all_symb_bigrams.write_string(order[pair[elem_num]], 0, chr(pair[elem_num]))
            all_symb_bigrams.write_number(order[pair[0]], order[pair[1]], pair[2], self.f_int)
        f_cond_rules = {'type': 'top', 'value': 10, 'criteria': '%', 'format': self.f_red_bg}
        all_symb_bigrams.conditional_format(1, 1, len(order), len(order), f_cond_rules)

    def sheet_top_words(self, limit=0, min_quantity=1):
        '''Create top-list of words by quantity. Is called from main "treat()".'''
        top_words = self.workbook.add_worksheet('Top words')
        top_words.set_tab_color('yellow')
        self.__add_main_style(top_words, 16)
        top_words.write_row(0, 0, ('Word', 'Quantity', '% from all', 'As first', 'As last'))
        if self.pos_list[2] and self.pos_list[2] != 1:
            top_words.write(0, 5, 'Avg. position')

        self.cursor.execute(
            f'''
            SELECT *
            FROM words
            WHERE quantity >= {min_quantity}
            ORDER BY quantity DESC, word ASC
            {f'LIMIT {limit}' if limit else ''};
            '''
        )
        for row, word in enumerate(self.cursor.fetchall(), 1):
            top_words.write_string(row, 0, word[0])
            top_words.write_number(row, 1, word[1], self.f_int)
            top_words.write_number(row, 2, word[1] / self.sum_list[2], self.f_percent)
            top_words.write_number(row, 3, word[2], self.f_int)
            top_words.write_number(row, 4, word[3], self.f_int)
            if self.pos_list[2]:
                top_words.write_number(row, 5, word[4], self.f_float)

    def sheet_word_bigrams_top(self, limit=0, min_quantity=1):
        '''Create top-list of word bigrams by quantity. Is called from main "treat()".'''
        word_bigrams_top = self.workbook.add_worksheet('Word bigrams top')
        word_bigrams_top.set_tab_color('yellow')
        self.__add_main_style(word_bigrams_top, 16, 12, True)
        word_bigrams_top.write_row(0, 0, ('First word', 'Second word', 'Quantity', '% from all'))
        if self.pos_list[3]:
            word_bigrams_top.write(0, 4, 'Avg. position')

        self.cursor.execute(
            f'''
            SELECT *
            FROM word_bigrams
            WHERE quantity >= {min_quantity}
            ORDER BY quantity DESC, first_word ASC, second_word ASC
            {f'LIMIT {limit}' if limit else ''};
            '''
        )
        for row, bigr in enumerate(self.cursor.fetchall(), 1):
            word_bigrams_top.write_string(row, 0, bigr[0])
            word_bigrams_top.write_string(row, 1, bigr[1])
            word_bigrams_top.write_number(row, 2, bigr[2], self.f_int)
            word_bigrams_top.write_number(row, 3, bigr[2] / self.sum_list[3], self.f_percent)
            if self.pos_list[3]:
                word_bigrams_top.write_number(row, 4, bigr[3], self.f_float)

    def sheet_custom_symb(self, symbols):
        '''Create symbol top-list with user inputed symbols.

        !This function is not called from main "treat()"
        '''
        symbols = {ord(x) for x in symbols}
        custom_top_symb = self.workbook.add_worksheet('Custom symbols')
        custom_top_symb.set_tab_color('blue')
        self.__add_main_style(custom_top_symb)
        custom_top_symb.write_row(0, 0, ('Symb', 'Quantity', '%', 'As first', 'As last'))
        if self.pos_list[0] != 1:
            custom_top_symb.write(0, 5, 'Avg. position')

        self.cursor.execute(
            f'''
            SELECT *
            FROM symbols
            WHERE ord IN {str(tuple(symbols))}
            ORDER BY quantity DESC, ord ASC;
            '''
        )

        for row, symb in enumerate(self.cursor.fetchall(), 1):
            custom_top_symb.write_string(row, 0, chr(symb[0]))
            custom_top_symb.write_number(row, 1, symb[1], self.f_int)
            custom_top_symb.write_formula(row, 2, f'=B{row+1}/SUM(B:B)', self.f_percent)
            custom_top_symb.write_number(row, 3, symb[2], self.f_int)
            custom_top_symb.write_number(row, 4, symb[3], self.f_int)
            if self.pos_list[0] != 1:
                custom_top_symb.write_number(row, 5, symb[4], self.f_float)

        chart = self.workbook.add_chart({'type': 'pie'})
        chart.add_series(
            {
                'name': 'Letter frequency',
                'categories': f'=\'Custom symbols\'!$A2:$A{len(symbols)+1}',
                'values': f'=\'Custom symbols\'!$C2:$C{len(symbols)+1}',
            }
        )
        custom_top_symb.insert_chart('H2', chart)
        print('... custom symbols top sheet was written.')

    def sheet_en_symb_bigrams(self):
        '''Create two-dimensional bigrams table only for English alphabet symbols.

        !This function is not called from main "treat()"!
        '''
        en_symb_bigrams = self.workbook.add_worksheet('English letter bigrams')
        en_symb_bigrams.set_tab_color('red')
        self.__add_main_style(en_symb_bigrams, 2.14, 9.43)
        order: dict = {
            **{x: n for n, x in enumerate(range(65, 91))},
            **{x: n for n, x in enumerate(range(97, 123))},
        }

        self.cursor.execute(
            '''
            SELECT *
            FROM symbol_bigrams
            WHERE (first_symb_ord BETWEEN 65 AND 90 OR first_symb_ord BETWEEN 97 AND 122)
                AND (second_symb_ord BETWEEN 65 AND 90 OR second_symb_ord BETWEEN 97 AND 122);
            '''
        )

        values: list = [[]]
        for pair in self.cursor.fetchall():
            loc_1 = order[pair[0]]
            loc_2 = order[pair[1]]
            while len(values) <= loc_1:
                values.append([])
            while len(values[loc_1]) <= loc_2:
                values[loc_1].append(0)
            values[loc_1][loc_2] += pair[2]

        en_symb_bigrams.write_row(0, 1, ascii_lowercase)
        en_symb_bigrams.write_column(1, 0, ascii_lowercase)

        for row, row_values_list in enumerate(values, 1):
            en_symb_bigrams.write_row(row, 1, row_values_list, self.f_int)
        f_cond_rules = {'type': 'top', 'value': 10, 'criteria': '%', 'format': self.f_red_bg}
        en_symb_bigrams.conditional_format(1, 1, 27, 27, f_cond_rules)
        print('... English letter bigrams sheet was written.')

    def sheet_ru_symb_bigrams(self):
        '''Create two-dimensional bigrams table only for Russian alphabet symbols.

        !This function is not called from main "treat()"!
        '''
        ru_symb_bigrams = self.workbook.add_worksheet('Russian letter bigrams')
        ru_symb_bigrams.set_tab_color('red')
        self.__add_main_style(ru_symb_bigrams, 2.14, 9.43)
        order: dict = {
            **{x: n for n, x in enumerate(range(1040, 1046), 0)},
            **{x: n for n, x in enumerate(range(1072, 1078), 0)},
            **{x: n for n, x in enumerate(range(1046, 1072), 7)},
            **{x: n for n, x in enumerate(range(1078, 1104), 7)},
            **{1025: 6, 1105: 6},
        }

        self.cursor.execute(
            '''
            SELECT *
            FROM symbol_bigrams
            WHERE (first_symb_ord BETWEEN 1040 AND 1105 OR first_symb_ord = 1025)
                AND (second_symb_ord BETWEEN 1040 AND 1105 OR second_symb_ord = 1025);
            '''
        )

        values: list = [[]]
        for pair in self.cursor.fetchall():
            loc_1 = order[pair[0]]
            loc_2 = order[pair[1]]
            while len(values) <= loc_1:
                values.append([])
            while len(values[loc_1]) <= loc_2:
                values[loc_1].append(0)
            values[loc_1][loc_2] += pair[2]

        rus_let = 'абвгдеёжзийклмнопрстуфхцчшщьыъэюя'
        ru_symb_bigrams.write_row(0, 1, rus_let)
        ru_symb_bigrams.write_column(1, 0, rus_let)

        for row, row_values_list in enumerate(values, 1):
            ru_symb_bigrams.write_row(row, 1, row_values_list, self.f_int)
        f_cond_rules = {'type': 'top', 'value': 10, 'criteria': '%', 'format': self.f_red_bg}
        ru_symb_bigrams.conditional_format(1, 1, 34, 34, f_cond_rules)
        print('... Russian letter bigrams sheet was written.')

    def sheet_yo_words(self, limit=0, min_quantity=1):
        '''Create sheet with quantity of entries for both of ye/yo word writing.

        !This function is not called from main "treat()"!
        '''
        yo_words = self.workbook.add_worksheet('Ye-yo words')
        yo_words.set_tab_color('yellow')
        self.__add_main_style(yo_words, 15, 15, True)
        yo_words.write_row(
            0, 0, ('Ё вариант', 'Е вариант', 'Ё обязательна?', 'Количество с Ё', 'Количество с Е')
        )

        counter = [0] * 3
        self.cursor.execute(
            '''
            SELECT SUM(quantity)
            FROM yo_words
            INNER JOIN words ON ye_word = word
            WHERE mandatory = 1;
            '''
        )
        counter[0] = self.cursor.fetchone()[0]
        self.cursor.execute(
            '''
            SELECT SUM(quantity)
            FROM yo_words
            INNER JOIN words ON ye_word = word
            WHERE mandatory = 0;
            '''
        )
        counter[1] = self.cursor.fetchone()[0]
        self.cursor.execute(
            '''
            SELECT SUM(quantity)
            FROM yo_words
            INNER JOIN words ON yo_word = word
            WHERE mandatory = 1;
            '''
        )
        counter[2] = self.cursor.fetchone()[0]

        self.cursor.execute(
            f'''
            SELECT
                yo_word, b.ye_word, mandatory,
                words.quantity as yo_quantity, ye_quantity
            FROM yo_words
            INNER JOIN words ON yo_word = word
            LEFT JOIN (
                SELECT ye_word, quantity as ye_quantity
                FROM yo_words
                INNER JOIN words ON ye_word = word
            ) b ON b.ye_word=yo_words.ye_word
            WHERE (yo_quantity + ye_quantity) >= {min_quantity}
            ORDER BY (yo_quantity + ye_quantity) DESC
            {f'LIMIT {limit}' if limit else ''};
            '''
        )

        for row, pair in enumerate(self.cursor.fetchall(), 1):
            yo_words.write_string(row, 0, pair[0])
            yo_words.write_string(row, 1, pair[1])
            yo_words.write_string(row, 2, ('Да' if pair[2] else 'Возможна'), self.f_bold)
            yo_words.write_number(row, 3, pair[3], self.f_int)
            yo_words.write_number(row, 4, pair[4], self.f_int)
        yo_words.write_string('G2', 'Ошибочная Е', self.f_bold)
        yo_words.write_number('H2', counter[0])
        yo_words.write_string('G3', 'Возможная Ё', self.f_bold)
        yo_words.write_number('H3', counter[1])
        yo_words.write_string('G4', 'Правильная Ё', self.f_bold)
        yo_words.write_number('H4', counter[2])

        print('... Russian ye/yo words compare sheet was written.')


class Result:
    '''Context manager with data validation for end-user ExcelWriter class.'''

    def __init__(self, name='frequency_analysis'):
        self.name = name
        self.db = None
        self.workbook = None

    def __enter__(self):
        if not re.search('^[a-zа-яё0-9_.@() -]+$', self.name, re.I):
            raise Exception(f"Foldername '{self.name}' is unvalid. Please, enter other.")
        if not os.path.isfile(os.path.join(os.getcwd(), self.name, 'result.db')):
            raise Exception(
                f"DB file in the '{self.name}' folder is not exist! "
                "Create a new analysis, or set name of folder with existing DB."
            )
        if os.path.isfile(os.path.join(os.getcwd(), self.name, 'result.xlsx')):
            raise Exception(
                f"xlsx file in the '{self.name}' folder already exist! "
                "Please, rename or delete an existing file."
            )

        self.db = sqlite3.connect(os.path.join(os.getcwd(), self.name, 'result.db'))
        self.workbook = xlsxwriter.Workbook(os.path.join(os.getcwd(), self.name, 'result.xlsx'))

        return ExcelWriter(self.workbook, self.db.cursor())

    def __exit__(self, type_, value, traceback):
        self.workbook.close()
        self.db.close()


__all__ = ['Result']
