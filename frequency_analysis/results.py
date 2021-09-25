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

        self.pos_list = [
            self.cursor.execute(f'SELECT AVG(position) FROM {x};').fetchone()[0]
            for x in ('symbols', 'symbol_bigrams', 'words', 'word_bigrams')
        ]
        self.count_list = [
            self.cursor.execute(f'SELECT COUNT(*) FROM {x};').fetchone()[0]
            for x in ('symbols', 'symbol_bigrams', 'words', 'word_bigrams')
        ]
        self.sum_list = [
            self.cursor.execute(f'SELECT COUNT(*) FROM {x};').fetchone()[0]
            for x in ('symbols', 'symbol_bigrams', 'words', 'word_bigrams')
        ]
        self.pos = lambda x: self.cursor.execute(f'SELECT AVG(position) FROM {x};').fetchone()[0]
        self.count = lambda x: self.cursor.execute(f'SELECT COUNT(*) FROM {x};').fetchone()[0]
        self.sum = lambda x: self.cursor.execute(f'SELECT SUM(quantity) FROM {x};').fetchone()[0]

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
            limits = tuple(list(limits) + [0] * (4 - len(limits)))
        if len(min_quantity) < 5:
            min_quantity = tuple(list(min_quantity) + [1] * (5 - len(min_quantity)))
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
            '(e.g. "sheet_en_symb_bigrams()", "sheet_ru_symb_bigrams()" or "yo_words()")'
        )

    def sheet_stats(self):
        '''Create main statistic of analysis. Is called from main "treat()".'''
        stats = self.workbook.add_worksheet('Stats')
        stats.write(0, 1, 'Total')
        stats.write(0, 2, 'Count')
        stats.write_column(1, 0, ('Symbols', 'Symbol bigrams', 'Words', 'Word bigrams'))
        stats.write_column(1, 1, self.count_list)
        stats.write_column(1, 2, self.sum_list)

    def sheet_symbols(self, limit=0, min_quantity=1):
        '''Create top-list of all analyzed symbols by quantity. Is called from main "treat()".'''
        symbols = self.workbook.add_worksheet('Symbols')
        symbols.freeze_panes(1, 1)
        symbols.write_row(0, 0, ('Symbol', 'Quantity', '% from all', 'As first', 'As last'))
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
            symbols.write_row(
                row, 0, (chr(symb[0]), symb[1], symb[1] / self.sum_list[0], symb[2], symb[3])
            )
            if self.pos_list != 1:
                symbols.write(row, 5, symb[4])

        chart = self.workbook.add_chart({'type': 'pie'})
        chart.add_series(
            {
                'name': 'Letter frequency',
                'categories': f'=Symbols!$A3:$A{limit+3 if limit else ""}',
                'values': f'=Symbols!$C3:$C{limit+3 if limit else ""}',
            }
        )
        symbols.insert_chart('H2', chart, {'x_offset': 25, 'y_offset': 10})

    def sheet_symbol_bigrams_top(self, limit=0, min_quantity=1):
        '''Create top-list of symbol bigrams by quantity. Is called from main "treat()".'''
        symbol_bigrams_top = self.workbook.add_worksheet('Symbol bigrams top')
        symbol_bigrams_top.freeze_panes(1, 1)
        symbol_bigrams_top.write_row(
            0, 0, ('First symbol', 'Second symbol', 'Quantity', '% from all')
        )
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
            symbol_bigrams_top.write_row(
                row, 0, (chr(bigr[0]), chr(bigr[1]), bigr[2], bigr[2] / self.sum_list[1])
            )
            if self.pos_list[1]:
                symbol_bigrams_top.write(row, 4, bigr[3])

    def sheet_all_symb_bigrams(self, min_quantity=1):
        '''Create 2D bigrams table for all analyzed symbols. Is called from main "treat()".'''
        all_symb_bigrams = self.workbook.add_worksheet('All symb bigrams')
        all_symb_bigrams.freeze_panes(1, 1)
        self.cursor.execute(
            f'''
            SELECT *
            FROM symbol_bigrams
            WHERE quantity >= {min_quantity}
            ORDER BY first_symb_ord ASC, second_symb_ord ASC;
            '''
        )
        locations: dict = {}
        for pair in self.cursor.fetchall():
            for elem_num in [0, 1]:
                if pair[elem_num] not in locations:
                    locations[pair[elem_num]] = len(locations) + 1
                    all_symb_bigrams.write(0, locations[pair[elem_num]], chr(pair[elem_num]))
                    all_symb_bigrams.write(locations[pair[elem_num]], 0, chr(pair[elem_num]))
            all_symb_bigrams.write(locations[pair[0]], locations[pair[1]], pair[2])

    def sheet_top_words(self, limit=0, min_quantity=1):
        '''Create top-list of words by quantity. Is called from main "treat()".'''
        top_words = self.workbook.add_worksheet('Top words')
        top_words.freeze_panes(1, 1)
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
            top_words.write_row(
                row, 0, (word[0], word[1], word[1] / self.sum_list[2], word[2], word[3])
            )
            if self.pos_list[2]:
                top_words.write(row, 5, word[4])

    def sheet_word_bigrams_top(self, limit=0, min_quantity=1):
        '''Create top-list of word bigrams by quantity. Is called from main "treat()".'''
        word_bigrams_top = self.workbook.add_worksheet('Word bigrams top')
        word_bigrams_top.freeze_panes(1, 1)
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
            word_bigrams_top.write_row(row, 0, (*bigr[:3], bigr[2] / self.sum_list[3]))
            if self.pos_list[3]:
                word_bigrams_top.write(row, 4, bigr[3])

    def sheet_en_symb_bigrams(self):
        '''Create two-dimensional bigrams table only for English alphabet symbols.

        !This function is not called from main "treat()"!
        '''
        en_symb_bigrams = self.workbook.add_worksheet('English letter bigrams')
        en_symb_bigrams.freeze_panes(1, 1)
        locations: dict = {
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
            loc_1 = locations[pair[1]]
            loc_2 = locations[pair[2]]
            while len(values) <= loc_1:
                values.append([])
            while len(values[loc_1]) <= loc_2:
                values[loc_1].append(0)
            values[loc_1][loc_2] += pair[3]

        en_symb_bigrams.write_row(0, 1, ascii_lowercase)
        en_symb_bigrams.write_column(1, 0, ascii_lowercase)

        for row, row_values_list in enumerate(values, 1):
            en_symb_bigrams.write_row(row, 1, row_values_list)

    def sheet_ru_symb_bigrams(self):
        '''Create two-dimensional bigrams table only for Russian alphabet symbols.

        !This function is not called from main "treat()"!
        '''
        ru_symb_bigrams = self.workbook.add_worksheet('Russian letter bigrams')
        ru_symb_bigrams.freeze_panes(1, 1)
        locations: dict = {
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
            loc_1 = locations[pair[1]]
            loc_2 = locations[pair[2]]
            while len(values) <= loc_1:
                values.append([])
            while len(values[loc_1]) <= loc_2:
                values[loc_1].append(0)
            values[loc_1][loc_2] += pair[3]

        rus_let = 'абвгдеёжзийклмнопрстуфхцчшщьыъэюя'
        ru_symb_bigrams.write_row(0, 1, rus_let)
        ru_symb_bigrams.write_column(1, 0, rus_let)

        for row, row_values_list in enumerate(values, 1):
            ru_symb_bigrams.write_row(row, 1, row_values_list)

    def sheet_yo_words(self, limit=0, min_quantity=1):
        '''Create sheet with quantity of entries for both of ye/yo word writing.

        !This function is not called from main "treat()"!
        '''
        yo_words = self.workbook.add_worksheet('Ye/yo words')
        yo_words.freeze_panes(1, 1)
        yo_words.write_row(0, 0, ('Yo word', 'Ye word', 'Yo mandatory?', 'Yo count', 'Ye count'))

        self.cursor.execute(
            f'''
            SELECT
                yo_word, ye_word, mandatory,
                word.quantity as yo_quantity,
                word.quantity as ye_quantity,
                SUM(yo_quantity, ye_quantity) as sum
            FROM yo_words
                INNER JOIN words ON yo_words.yo_word = words.word
                INNER JOIN words ON yo_words.ye_word = words.word
            WHERE sum >= {min_quantity}
            ORDER BY sum
            {f'LIMIT {limit}' if limit else ''};
            '''
        )

        for row, pair in enumerate(self.cursor.fetchall(), 1):
            yo_words.write(row, 0, pair[0])
            yo_words.write(row, 1, pair[1])
            yo_words.write(row, 2, ('Mandatory' if pair[2] else 'Probably'))
            self.cursor.execute(f"SELECT quantity FROM words WHERE word='{pair[0]}';")
            yo_words.write(row, 3, self.cursor.fetchone()[0])
            self.cursor.execute(f"SELECT quantity FROM words WHERE word='{pair[1]}';")
            yo_words.write(row, 4, self.cursor.fetchone()[0])


class Result:
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

        self.db = sqlite3.connect(os.path.join(os.getcwd(), self.name, 'result.db'))
        self.workbook = xlsxwriter.Workbook(os.path.join(os.getcwd(), self.name, 'result.xlsx'))

        return ExcelWriter(self.workbook, self.db.cursor())

    def __exit__(self, type_, value, traceback):
        self.workbook.close()
        self.db.close()


__all__ = ['Result']
