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

    def __add_main_style(
        self, sheet, f_width=5, a_width=12, two_columns=False, two_rows=0, color=None
    ):
        '''Add main row/column formating (width, bold, centred, freeze).'''
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

    def sheet_symbols(self, limit=0, min_quantity=1):
        '''Create top-list of all analyzed symbols by quantity. Is called from main "treat()".'''
        symbols = self.workbook.add_worksheet('Symbols')
        self.__add_main_style(symbols, color='green', two_rows=6)

        self.cursor.execute('SELECT * FROM symbols')
        values: list = [{}, {}]  # [case sensitive, case insensitive]
        for symb in self.cursor.fetchall():
            values[0][symb[0]] = list(symb[1:])
            if (s := symb[0].lower()) in values[1]:
                if self.pos_list[0] != 1:
                    values[1][s][3] = (values[1][s][3] * values[1][s][0] + symb[4] * symb[1]) / (
                        values[1][s][0] + symb[1]
                    )
                values[1][s][0] += symb[1]
                values[1][s][1] += symb[2]
                values[1][s][2] += symb[3]
            else:
                values[1][s] = list(symb[1:])

        for e in [0, 7]:
            symbols.write_row(1, e, ('Symb', 'Quantity', '% from all', 'As first', 'As last'))
            if self.pos_list[0] != 1:
                symbols.write(1, 5 + e, 'Avg. position')
            values[bool(e)] = dict(
                sorted(values[bool(e)].items(), key=lambda x: x[1][0], reverse=True)
            )
            for row, (symb, vals) in enumerate(values[bool(e)].items(), 2):
                if (limit and row > limit + 1) or vals[0] < min_quantity:
                    break
                symbols.write_string(row, 0 + e, symb)
                symbols.write_number(row, 1 + e, vals[0], self.f_int)
                symbols.write_number(row, 2 + e, vals[0] / self.sum_list[0], self.f_percent)
                if symb != ' ':
                    symbols.write_number(row, 3 + e, vals[1], self.f_int)
                    symbols.write_number(row, 4 + e, vals[2], self.f_int)
                    if self.pos_list[0] != 1:
                        symbols.write_number(row, 5 + e, vals[3], self.f_float)
            chart = self.workbook.add_chart({'type': 'pie'})
            cols = ('H', 'J') if bool(e) else ('A', 'C')
            chart.add_series(
                {
                    'name': f'Letter frequency (case {"in" if bool(e) else ""}sensitive)',
                    'categories': f'=Symbols!${cols[0]}3:${cols[0]}{limit+3 if limit else ""}',
                    'values': f'=Symbols!${cols[1]}3:${cols[1]}{limit+3 if limit else ""}',
                }
            )
            symbols.insert_chart(f'O{18 if bool(e) else 3}', chart)

    def sheet_symbol_bigrams_top(self, limit=0, min_quantity=1):
        '''Create top-list of symbol bigrams by quantity. Is called from main "treat()".'''
        symbol_bigrams_top = self.workbook.add_worksheet('Symbol bigrams top')
        self.__add_main_style(symbol_bigrams_top, two_columns=True, two_rows=7, color='green')

        self.cursor.execute('SELECT * FROM symbol_bigrams')

        values: list = [{}, {}]
        for symb in self.cursor.fetchall():
            values[0][symb[0] + symb[1]] = list(symb[2:])
            if (s := symb[0].lower() + symb[1].lower()) in values[1]:
                if self.pos_list[1]:
                    values[1][s][3] = (values[1][s][3] * values[1][s][0] + symb[5] * symb[2]) / (
                        values[1][s][0] + symb[2]
                    )
                values[1][s][0] += symb[2]
                values[1][s][1] += symb[3]
                values[1][s][2] += symb[4]
            else:
                values[1][s] = list(symb[2:])
        for e in [0, 8]:
            symbol_bigrams_top.write_row(
                1, e, ('1st', '2nd', 'Quantity', '% from all', 'As first', 'As last')
            )
            if self.pos_list[1]:
                symbol_bigrams_top.write(1, 6 + e, 'Avg. position')
            values[bool(e)] = dict(
                sorted(values[bool(e)].items(), key=lambda x: x[1][0], reverse=True)
            )
            for row, (pair, vals) in enumerate(values[bool(e)].items(), 2):
                if (limit and row > limit + 1) or vals[0] < min_quantity:
                    break
                symbol_bigrams_top.write_string(row, 0 + e, pair[0])
                symbol_bigrams_top.write_string(row, 1 + e, pair[1])
                symbol_bigrams_top.write_number(row, 2 + e, vals[0], self.f_int)
                symbol_bigrams_top.write_number(
                    row, 3 + e, vals[0] / self.sum_list[1], self.f_percent
                )
                symbol_bigrams_top.write_number(row, 4 + e, vals[1], self.f_int)
                symbol_bigrams_top.write_number(row, 5 + e, vals[2], self.f_int)
                if self.pos_list[1]:
                    symbol_bigrams_top.write_number(row, 6 + e, vals[3], self.f_float)
            chart = self.workbook.add_chart({'type': 'pie'})
            cols = ('I', 'J', 'L') if bool(e) else ('A', 'B', 'D')
            chart.add_series(
                {
                    'name': f'Symbol bigrams frequency (case {"in" if bool(e) else ""}sensitive)',
                    'categories': (
                        f'=\'Symbol bigrams top\'!${cols[0]}3:${cols[1]}{limit+3 if limit else ""}'
                    ),
                    'values': (
                        f'=\'Symbol bigrams top\'!${cols[2]}3:${cols[2]}{limit+3 if limit else ""}'
                    ),
                }
            )
            symbol_bigrams_top.insert_chart(f'Q{18 if bool(e) else 3}', chart)

    def sheet_all_symb_bigrams(self, min_quantity=1, ignore_case=False):
        '''Create 2D bigrams table for all analyzed symbols. Is called from main "treat()".'''
        all_symb_bigrams = self.workbook.add_worksheet(
            f'All symb bigrams{" [I]" if ignore_case else ""}'
        )
        all_symb_bigrams.set_tab_color('red')
        self.__add_main_style(all_symb_bigrams, 2.14, 9.43)
        self.cursor.execute('SELECT DISTINCT first_symb FROM symbol_bigrams;')
        fst_symbs = [x[0] for x in self.cursor.fetchall()]
        self.cursor.execute('SELECT DISTINCT second_symb FROM symbol_bigrams;')
        snd_symbs = [x[0] for x in self.cursor.fetchall()]
        all_symbs = set(fst_symbs + snd_symbs)
        order: dict = {}
        values: dict = {}
        for symb in all_symbs.copy():
            self.cursor.execute(
                f'''
                SELECT SUM(quantity)
                FROM symbol_bigrams
                WHERE first_symb='{symb}' OR second_symb='{symb}'
                {' COLLATE NOCASE' if ignore_case else ''};
                '''
            )
            if (s := self.cursor.fetchone()[0]) and s >= min_quantity:
                self.cursor.execute(
                    f'''
                    SELECT *
                    FROM symbol_bigrams
                    WHERE first_symb='{symb}' OR second_symb='{symb}'
                    {' COLLATE NOCASE' if ignore_case else ''};
                    '''
                )
                for pair in self.cursor.fetchall():
                    for e in [0, 1]:
                        if (s := pair[e].lower() if ignore_case else pair[e]) not in order:
                            order[s] = len(order) + 1
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
        for bigr, val in values.items():
            all_symb_bigrams.write_number(order[bigr[0]], order[bigr[1]], val[0], self.f_int)
            all_symb_bigrams.write_comment(
                order[bigr[0]],
                order[bigr[1]],
                f'As first: {val[1]}; as last: {val[2]}' + f'; position: {val[3]}'
                if self.pos_list[1]
                else '',
            )
        f_cond_rules = {'type': 'top', 'value': 10, 'criteria': '%', 'format': self.f_red_bg}
        all_symb_bigrams.conditional_format(1, 1, len(order) + 1, len(order) + 1, f_cond_rules)

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

    def sheet_custom_symb(self, symbols, name='Custom symbols'):
        '''Create symbol top-list with user inputed symbols.

        !This function is not called from main "treat()"
        '''
        i_symbols = {y for x in symbols for y in (ord(x.lower()), ord(x.upper()))}
        symbols = {ord(x) for x in symbols}
        while True:
            try:
                custom_top_symb = self.workbook.add_worksheet(name)
                break
            except xlsxwriter.exceptions.DuplicateWorksheetName:
                name += ' – Copy'
        custom_top_symb.set_tab_color('gray')
        self.__add_main_style(custom_top_symb)
        custom_top_symb.freeze_panes(2, 2)
        custom_top_symb.set_row(1, None, self.f_bold)
        custom_top_symb.set_column('G:H', 5, self.f_bold)
        custom_top_symb.set_column('B:F', 12)
        custom_top_symb.set_column('I:AZ', 12)
        custom_top_symb.merge_range('A1:F1', 'Case sensitive', self.f_bold)
        custom_top_symb.merge_range('H1:M1', 'Case insensitive', self.f_bold)
        custom_top_symb.write_row(1, 0, ('Symb', 'Quantity', '%', 'As first', 'As last'))
        custom_top_symb.write_row(1, 7, ('Symb', 'Quantity', '%', 'As first', 'As last'))
        if self.pos_list[0] != 1:
            custom_top_symb.write(1, 5, 'Avg. position')
            custom_top_symb.write(1, 12, 'Avg. position')

        self.cursor.execute(
            f'''
            SELECT *
            FROM symbols
            WHERE ord IN {str(tuple(symbols))}
            ORDER BY quantity DESC, ord ASC;
            '''
        )

        res = self.cursor.fetchall()
        values: dict = {}
        if ignore_case:
            for symb in res:
                if (s := chr(symb[0]).lower()) in values:
                    if self.pos_list[0] != 1:
                        values[s][3] = (values[s][3] * values[s][0] + symb[1] * symb[4]) / (
                            values[s][0] + symb[1]
                        )
                    values[s][0] += symb[1]
                    values[s][1] += symb[2]
                    values[s][2] += symb[3]
                else:
                    values[s] = list(symb[1:])
        else:
            for symb in self.cursor.fetchall():
                values[chr(symb[0])] = symb[1:]

        for row, (symb, value) in enumerate(values.items(), 1):
            custom_top_symb.write_string(row, 0, str(symb))
            custom_top_symb.write_number(row, 1, value[0], self.f_int)
            custom_top_symb.write_formula(row, 2, f'=B{row+1}/SUM(B:B)', self.f_percent)
            custom_top_symb.write_number(row, 3, value[1], self.f_int)
            custom_top_symb.write_number(row, 4, value[2], self.f_int)
            if self.pos_list[0] != 1:
                custom_top_symb.write_number(row, 5, value[3], self.f_float)

        for row, (symb, value) in enumerate(values.items(), 1):
            custom_top_symb.write_string(row, 0, str(symb))
            custom_top_symb.write_number(row, 1, value[0], self.f_int)
            custom_top_symb.write_formula(row, 2, f'=B{row+1}/SUM(B:B)', self.f_percent)
            custom_top_symb.write_number(row, 3, value[1], self.f_int)
            custom_top_symb.write_number(row, 4, value[2], self.f_int)
            if self.pos_list[0] != 1:
                custom_top_symb.write_number(row, 5, value[3], self.f_float)

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
