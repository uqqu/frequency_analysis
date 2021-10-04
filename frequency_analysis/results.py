'''Additional module to frequency.py for excel output.'''

import os
import re
import sqlite3
from ast import literal_eval
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
        self.f_int_null = self.workbook.add_format({'align': 'center', 'color': 'gray'})
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

    def __add_main_style(
        self, sheet, f_width=5, a_width=12, *, two_columns=False, two_rows=0, color=None
    ):
        '''Add main row/column formating (width, bold, centred, freeze, color, merge, errors).'''
        if color:
            sheet.set_tab_color(color)
        sheet.ignore_errors({'number_stored_as_text': 'A:ZZ'})
        sheet.freeze_panes(1 + bool(two_rows), 1 + two_columns)
        sheet.set_row(0, None, self.f_bold)
        sheet.set_column(0, 0 + two_columns, f_width, self.f_bold)
        if two_rows:
            sheet.set_column(two_rows + 1, two_rows + 1 + two_columns, f_width, self.f_bold)
            sheet.set_column(1 + two_columns, two_rows, a_width)
            sheet.set_column(two_rows + 2 + two_columns, 99, a_width)
            sheet.set_row(1, None, self.f_bold)
            sheet.merge_range(0, 0, 0, two_rows - 1, 'Case sensitive', self.f_bold)
            sheet.merge_range(0, two_rows + 1, 0, two_rows * 2, 'Case insensitive', self.f_bold)
        else:
            sheet.set_column(1 + two_columns, 99, a_width)

    def __fill_top_data(
        self,
        sheet,
        table_name: str,
        pos_data: bool,
        title_data: tuple,
        dbl: bool,
        limit: int,
        min_quantity: int,
        chart_limit: int,
        sum_value: int,
    ):
        '''Fill data for 1D (top-list) sheets.'''
        values: list = [{}, {}]  # [case sensitive, case insensitive]
        for symb in self.cursor.fetchall():
            values[0][symb[0] + (symb[1] if dbl else '')] = list(symb[1 + dbl :])
            if (s := (symb[0] + (symb[1] if dbl else '')).lower()) in values[1]:
                if pos_data:
                    values[1][s][3] = (
                        values[1][s][3] * values[1][s][0] + symb[4 + dbl] * symb[1 + dbl]
                    ) / (values[1][s][0] + symb[1 + dbl])
                values[1][s][0] += symb[1 + dbl]
                values[1][s][1] += symb[2 + dbl]
                values[1][s][2] += symb[3 + dbl]
            else:
                values[1][s] = list(symb[1 + dbl :])

        for e in [0, 6 + len(title_data)]:
            sheet.write_row(1, e, title_data + ('Quantity', '% from all', 'As first', 'As last'))
            if pos_data:
                sheet.write(1, 4 + len(title_data) + e, 'Avg. position')
            values[bool(e)] = dict(
                sorted(values[bool(e)].items(), key=lambda x: x[1][0], reverse=True)
            )
            for row, (symb, vals) in enumerate(values[bool(e)].items(), 2):
                if (limit and row > limit + 1) or vals[0] < min_quantity:
                    break
                sheet.write_string(row, 0 + e, symb[0])
                if dbl:
                    sheet.write_string(row, 1 + e, symb[1])
                sheet.write_number(row, 1 + dbl + e, vals[0], self.f_int)
                if sum_value:
                    sheet.write_number(row, 2 + dbl + e, vals[0] / sum_value, self.f_percent)
                else:
                    c = chr(66 + dbl + e)
                    sheet.write_formula(
                        row, 2 + dbl + e, f'={c}{row + 1}/SUM({c}:{c})', self.f_percent
                    )
                if symb != ' ':
                    sheet.write_number(row, 3 + dbl + e, vals[1], self.f_int)
                    sheet.write_number(row, 4 + dbl + e, vals[2], self.f_int)
                    if pos_data:
                        sheet.write_number(row, 5 + dbl + e, vals[3], self.f_float)
            chart = self.workbook.add_chart({'type': 'pie'})
            chart.add_series(
                {
                    'name': f'Case {"in" if bool(e) else ""}sensitive | Top {chart_limit}',
                    'categories': f"='{table_name}'!${chr(65+e)}3:${chr(65+e+dbl)}{chart_limit+2}",
                    'values': f"='{table_name}'!${chr(67+e+dbl)}3:${chr(67+e+dbl)}{chart_limit+2}",
                }
            )
            chart.set_size({'width': 356, 'height': 360})
            chart.set_legend({'layout': {'x': 0.95, 'y': 0.37, 'width': 0.13, 'height': 0.95}})
            chart.set_style(6)
            sheet.insert_chart(f'{"Q" if dbl else "O"}{21 if bool(e) else 3}', chart)

    def __2d_symb_bigrams(self, sheet, min_quantity: int, ignore_case: bool, custom_symbols=''):
        '''Fill data for 2D (bigrams n:n) sheets.'''
        self.cursor.execute('SELECT DISTINCT first_symb FROM symbol_bigrams;')
        fst_symbs = [x[0] for x in self.cursor.fetchall()]
        self.cursor.execute('SELECT DISTINCT second_symb FROM symbol_bigrams;')
        all_symbs = set(fst_symbs + [x[0] for x in self.cursor.fetchall()])
        if custom_symbols:
            all_symbs &= set(custom_symbols)
        if ignore_case:
            all_symbs = {x.lower() for x in all_symbs}
        order: list = []
        values: dict = {}
        for symb in all_symbs:
            clear_symb = symb.replace("'", "''")
            if ignore_case:
                where = f'''
                    WHERE first_symb='{clear_symb.lower()}' OR first_symb='{clear_symb.upper()}';
                    '''
            else:
                where = f"WHERE first_symb='{clear_symb}';"
            self.cursor.execute(f'SELECT SUM(quantity) FROM symbol_bigrams {where}')
            if (s := self.cursor.fetchone()[0]) and s >= min_quantity:
                order.append(symb)
                self.cursor.execute(f'SELECT * FROM symbol_bigrams {where}')
                for pair in self.cursor.fetchall():
                    b = (pair[0] + pair[1]).lower() if ignore_case else (pair[0] + pair[1])
                    if b in values:
                        if self.pos_list[1]:
                            values[b][3] = (values[b][0] * values[b][3] + pair[2] * pair[5]) / (
                                values[b][0] + pair[2]
                            )
                        values[b][0] += pair[2]
                        values[b][1] += pair[3]
                        values[b][2] += pair[4]
                    else:
                        values[b] = list(pair[2:])
        if custom_symbols:
            order = sorted(order, key=custom_symbols.index)
        else:
            order = sorted(order)
        for pos, symb in enumerate(order, 1):
            sheet.write_string(pos, 0, symb, self.f_bold)
            sheet.write_string(0, pos, symb, self.f_bold)
            sheet.write_row(pos, pos, (0,) * (len(order) - pos + 1), self.f_int_null)
            sheet.write_column(pos, pos, (0,) * (len(order) - pos + 1), self.f_int_null)
        for bigr, val in values.items():
            if bigr[0] not in order or bigr[1] not in order:
                continue
            sheet.write_number(
                order.index(bigr[0]) + 1, order.index(bigr[1]) + 1, val[0], self.f_int
            )
            sheet.write_comment(
                order.index(bigr[0]) + 1,
                order.index(bigr[1]) + 1,
                f'As first: {val[1]}; as last: {val[2]}' + f'; position: {round(val[3], 2)}'
                if self.pos_list[1]
                else '',
            )
        f_cond_rules = {'type': 'top', 'value': 10, 'criteria': '%', 'format': self.f_red_bg}
        sheet.conditional_format(1, 1, len(order) + 1, len(order) + 1, f_cond_rules)

    def treat(self, limits=(0,) * 4, min_quantity=(1,) * 5, chart_limit=(20,) * 4):
        '''Create main sheets all at once.

        Input:
            limits  — tuple – max number of elements to be added to the sheet (0 – unlimited))
                    — (symbols, symbol bigrams top, words top, word bigrams top)
                    — default values – [0, 0, 0, 0] (ommited = default);
            min_quantity – tuple – min number of entries for each element to take it into account)
                    — (/same as on 'limits'/, +symbol bigrams table)
                    — default values – (1, 1, 1, 1, 1) (ommited = default)
                    — (symbs, symb bigrs top, words top, word bigrs top, !symbol bigrams table!).
            chart_limit – tuple – number of first items for pie charts
                    — (/same as on 'limits'/)
                    — (symbols, symbol bigrams top, words top, word bigrams top)
        '''
        for t, (l, v) in {limits: (4, 0), min_quantity: (5, 1), chart_limit: (4, 20)}.items():
            if len(t) < l:
                literal_eval(f'{t} = tuple({t}) + ({v},) * ({l} - len({t}))')
        print('Start of writing to .xlsx')
        self.sheet_stats()
        print('... stats sheet was written')
        self.sheet_symbols(limits[0], min_quantity[0], chart_limit[0])
        print('... symbols sheet was written')
        self.sheet_symbol_bigrams_top(limits[1], min_quantity[1], chart_limit[1])
        print('... top symbol bigrams sheet was written')
        self.sheet_all_symb_bigrams(min_quantity[4])
        print('... symbol bigrams table sheet was written')
        self.sheet_top_words(limits[2], min_quantity[2], chart_limit[2])
        print('... top words sheet was written')
        self.sheet_word_bigrams_top(limits[3], min_quantity[3], chart_limit[3])
        print('... top word bigrams sheet was written')
        print('End of writing main sheets.')
        print(
            'You can call additional functions to create more sheets '
            '(e.g. "sheet_en_symb_bigrams()", "sheet_ru_symb_bigrams()", '
            '"sheet_yo_words([limit], [min_quantity])") or "sheet_custom_symb(symbols_string)".\n'
            'You also can call 2D sheet functions with "ignore_case=True" argument.'
        )

    def sheet_stats(self):
        '''Create main statistic of analysis. Is called from main "treat()".'''
        count_list = [
            self.cursor.execute(f'SELECT COUNT(*) FROM {x};').fetchone()[0]
            for x in ('symbols', 'symbol_bigrams', 'words', 'word_bigrams')
        ]
        avg_pos_list = [
            self.cursor.execute(
                f'''
                SELECT SUM(quantity*position)/SUM(quantity)
                FROM {x}
                {"WHERE chr != ' '" if x == "symbols" else ""};
                '''
            ).fetchone()[0]
            for x in ('symbols', 'symbol_bigrams', 'words', 'word_bigrams')
        ]
        stats = self.workbook.add_worksheet('Stats')
        self.__add_main_style(stats, 15, 15)
        stats.write_row(0, 1, ('Total', 'Quantity', 'Avg. position'))
        stats.write_column(1, 0, ('Symbols', 'Symbol bigrams', 'Words', 'Word bigrams'))
        stats.write_column(1, 1, count_list, self.f_int)
        stats.write_column(1, 2, self.sum_list, self.f_int)
        stats.write_column(1, 3, avg_pos_list, self.f_float)

    def sheet_symbols(self, limit=0, min_quantity=1, chart_limit=20):
        '''Create top-list of all analyzed symbols by quantity. Is called from main "treat()".'''
        symbols = self.workbook.add_worksheet('Symbols')
        self.__add_main_style(symbols, two_rows=6, color='green')
        self.cursor.execute('SELECT * FROM symbols')
        self.__fill_top_data(
            symbols,
            'Symbols',
            self.pos_list[0] != 1,
            ('Symb',),
            False,
            limit,
            min_quantity,
            chart_limit,
            self.sum_list[0],
        )

    def sheet_symbol_bigrams_top(self, limit=0, min_quantity=1, chart_limit=20):
        '''Create top-list of symbol bigrams by quantity. Is called from main "treat()".'''
        symbol_bigrams_top = self.workbook.add_worksheet('Symbol bigrams')
        self.__add_main_style(symbol_bigrams_top, two_columns=True, two_rows=7, color='green')
        self.cursor.execute('SELECT * FROM symbol_bigrams')
        self.__fill_top_data(
            symbol_bigrams_top,
            'Symbol bigrams',
            bool(self.pos_list[1]),
            ('1st', '2nd'),
            True,
            limit,
            min_quantity,
            chart_limit,
            self.sum_list[1],
        )

    def sheet_all_symb_bigrams(self, min_quantity=1, *, ignore_case=False):
        '''Create 2D bigrams for all analyzed symbols. Is called from main "treat()".'''
        all_symb_bigrams = self.workbook.add_worksheet(
            f'All symb bigrams{" (I)" if ignore_case else ""}'
        )
        self.__add_main_style(all_symb_bigrams, 2.14, 9.43)
        self.__2d_symb_bigrams(all_symb_bigrams, min_quantity, ignore_case)

    def sheet_top_words(self, limit=0, min_quantity=1, chart_limit=20):
        '''Create top-list of words by quantity. Is called from main "treat()".'''
        top_words = self.workbook.add_worksheet('Top words')
        self.__add_main_style(top_words, 16, color='yellow')
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

        chart = self.workbook.add_chart({'type': 'pie'})
        chart.add_series(
            {
                'name': f'Top {chart_limit}',
                'categories': f"='Top words'!$A2:$A{chart_limit+1}",
                'values': f"='Top words'!$C2:$C{chart_limit+1}",
            }
        )
        chart.set_size({'width': 356, 'height': 360})
        chart.set_legend({'layout': {'x': 0.95, 'y': 0.37, 'width': 0.12, 'height': 0.95}})
        chart.set_style(6)
        top_words.insert_chart('H2', chart)

    def sheet_word_bigrams_top(self, limit=0, min_quantity=1, chart_limit=20):
        '''Create top-list of word bigrams by quantity. Is called from main "treat()".'''
        word_bigrams_top = self.workbook.add_worksheet('Word bigrams top')
        self.__add_main_style(word_bigrams_top, 16, 12, two_columns=True, color='yellow')
        word_bigrams_top.write_row(
            0, 0, ('First word', 'Second word', 'Quantity', '% from all', 'As first', 'As last')
        )
        if self.pos_list[3]:
            word_bigrams_top.write(0, 6, 'Avg. position')

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
            word_bigrams_top.write_number(row, 4, bigr[3], self.f_int)
            word_bigrams_top.write_number(row, 5, bigr[4], self.f_int)
            if self.pos_list[3]:
                word_bigrams_top.write_number(row, 6, bigr[5], self.f_float)

        chart = self.workbook.add_chart({'type': 'pie'})
        chart.add_series(
            {
                'name': f'Top {chart_limit}',
                'categories': f"='Word bigrams top'!$A2:$B{chart_limit+1}",
                'values': f"='Word bigrams top'!$C2:$C{chart_limit+1}",
            }
        )
        chart.set_size({'width': 534, 'height': 360})
        chart.set_legend({'layout': {'x': 0.95, 'y': 0.37, 'width': 0.37, 'height': 0.95}})
        chart.set_style(6)
        word_bigrams_top.insert_chart('I2', chart)

    def sheet_custom_symb(self, symbols: str, chart_limit=20, *, name='Custom symbols'):
        '''Create symbol top-list with user inputed symbols.

        !This function is not called from main "treat()"
        '''
        while True:
            try:
                custom_top_symb = self.workbook.add_worksheet(name)
                break
            except xlsxwriter.exceptions.DuplicateWorksheetName:
                name += ' – Copy'
        self.__add_main_style(custom_top_symb, two_rows=6, color='gray')
        self.cursor.execute(f'SELECT * FROM symbols WHERE chr IN {str(tuple(symbols))}')
        self.__fill_top_data(
            custom_top_symb, name, self.pos_list[0], ('Symb',), False, 0, 0, chart_limit, 0
        )
        print('... custom symbols top sheet was written.')

    def sheet_en_symb_bigrams(self, *, ignore_case=False):
        '''Create two-dimensional bigrams table only for English alphabet symbols.

        !This function is not called from main "treat()"!
        '''
        en_symb_bigrams = self.workbook.add_worksheet(
            f'English letter bigrams{" (I)" if ignore_case else ""}'
        )
        self.__add_main_style(en_symb_bigrams, 2.14, 9.43)
        self.__2d_symb_bigrams(en_symb_bigrams, 1, ignore_case, ascii_lowercase)
        print('... English letter bigrams sheet was written.')

    def sheet_ru_symb_bigrams(self, *, ignore_case=False):
        '''Create two-dimensional bigrams table only for Russian alphabet symbols.

        !This function is not called from main "treat()"!
        '''
        ru_symb_bigrams = self.workbook.add_worksheet(
            f'Russian letter bigrams{" (I)" if ignore_case else ""}'
        )
        self.__add_main_style(ru_symb_bigrams, 2.14, 9.43)
        self.__2d_symb_bigrams(
            ru_symb_bigrams, 1, ignore_case, 'абвгдеёжзийклмнопрстуфхцчшщьыъэюя'
        )
        print('... Russian letter bigrams sheet was written.')

    def sheet_custom_symb_bigrams(
        self, symbols, *, ignore_case=False, name='Custom symbol bigrams'
    ):
        '''Create two-dimensional bigrams table only for English alphabet symbols.

        !This function is not called from main "treat()"!
        '''
        if ignore_case:
            name += ' (I)'
        while True:
            try:
                custom_symb_bigrams = self.workbook.add_worksheet(name)
                break
            except xlsxwriter.exceptions.DuplicateWorksheetName:
                name += ' – Copy'
        self.__add_main_style(custom_symb_bigrams, 2.14, 9.43)
        self.__2d_symb_bigrams(custom_symb_bigrams, 1, ignore_case, symbols)
        print('... custom symbol bigrams sheet was written.')

    def sheet_yo_words(self, limit=0, min_quantity=1):
        '''Create sheet with quantity of entries for both of ye/yo word writing.

        !This function is not called from main "treat()"!
        '''
        yo_words = self.workbook.add_worksheet('Ye-yo words')
        yo_words.set_tab_color('yellow')
        self.__add_main_style(yo_words, 15, 15, two_columns=True)
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
