import sqlite3
import xlsxwriter


def main(cursor, limits=(0, 0, 0, 0), min_quantity=(1, 1, 1, 1, 1), pos=False):
    ''' Call all mandatory functions.

    Additional functions – sheet_ru_symb_bigrams and yo_words_sheet are called individually.

    Input:
        cursor – sqlite db.cursor()
        limits – tuple (max number of elements to be added to the sheet (0 – unlimited))
            – (symbols, symbol bigrams top, words top, word bigrams top)
        min_quantity – tuple (min number of entries for each element to take it into account)
            – ([same as on 'limits'], symbol bigrams table)
            – (symbols, symb bigrams top, words top, word bigrams top, !symbol bigrams table!)
        pos – whether the position was taken into account in the analysis
    '''
    print('Start of writing to .xlsx')
    sheet_stats(cursor)
    print('')
    sheet_symbols(cursor, limits[0], min_quantity[0], pos)
    sheet_symbol_bigrams_top(cursor, limits[1], min_quantity[0], pos)



def sheet_stats(workbook, cursor):
    stats = workbook.add_worksheet('Stats')
    stats.write(0, 1, 'Total')
    stats.write(0, 2, 'Count')
    stats.write(1, 0, 'Symbols')
    stats.write(2, 0, 'Symbol bigrams')
    stats.write(3, 0, 'Words')
    stats.write(4, 0, 'Word bigrams')

    count = lambda x: cursor.execute(f'SELECT COUNT(*) FROM {x};')
    quantity = lambda x: cursor.execute(f'SELECT SUM(quantity) FROM {x};')

    count('symbols')
    stats.write(1, 1, cursor.fetchone()[0])
    quantity('symbols')
    stats.write(1, 2, cursor.fetchone()[0])
    count('symbol_bigrams')
    stats.write(2, 1, cursor.fetchone()[0])
    quantity('symbol_bigrams')
    stats.write(2, 2, cursor.fetchone()[0])
    count('words')
    stats.write(3, 1, cursor.fetchone()[0])
    quantity('words')
    stats.write(3, 2, cursor.fetchone()[0])
    count('word_bigrams')
    stats.write(4, 1, cursor.fetchone()[0])
    quantity('word_bigrams')
    stats.write(4, 2, cursor.fetchone()[0])


def sheet_symbols(cursor, limit=0, min_quantity=1, pos=False):
    symbols = workbook.add_worksheet('Symbols')
    symbols.write(0, 0, 'Symbol')
    symbols.write(0, 1, 'Quantity')
    symbols.write(0, 2, r'% from all')
    symbols.write(0, 3, 'As first')
    symbols.write(0, 4, 'As last')
    if pos:
        symbols.write(0, 5, 'Avg. position')

    cursor.execute('SELECT SUM(quantity) FROM symbols;')
    all_sum = cursor.fetchone()[0]

    cursor.execute(
        f'''
        SELECT *
        FROM symbols
        WHERE quantity >= {min_quantity}
        ORDER BY quantity DESC, ord ASC
        {f'LIMIT {limit}' if limit else ''};
        '''
    )

    for row, symb in enumerate(cursor.fetchall(), 1):
        symbols.write(row, 0, chr(symb[1]))
        symbols.write(row, 1, symb[2])
        symbols.write(row, 2, symb[2] / all_sum)
        symbols.write(row, 3, symb[3])
        symbols.write(row, 4, symb[4])
        if pos:
            cursor.execute(
                f'''
                SELECT AVG(position)
                FROM symbol_entries
                WHERE symb_id = {str(symb[0])};
                '''
            )
            symbols.write(row, 5, (cursor.fetchone()[0] or 0))


def sheet_symbol_bigrams_top(cursor, limit=0, min_quantity=1, pos=False):
    symbol_bigrams_top = workbook.add_worksheet('Symbol bigrams top')
    symbol_bigrams_top.write(0, 0, 'First symb')
    symbol_bigrams_top.write(0, 1, 'Second symb')
    symbol_bigrams_top.write(0, 2, 'Quantity')
    symbol_bigrams_top.write(0, 3, r'% from all')
    if pos:
        symbol_bigrams_top.write(0, 4, 'Avg. position')

    cursor.execute('SELECT SUM(quantity) FROM symbol_bigrams;')
    all_sum = cursor.fetchone()[0]

    cursor.execute(
        f'''
        SELECT *
        FROM symbol_bigrams
        WHERE quantity >= {min_quantity}
        ORDER BY quantity DESC, first_symb ASC, second_symb ASC
        {f'LIMIT {limit}' if limit else ''};
        '''
    )

    for row, bigr in enumerate(cursor.fetchall(), 1):
        symbol_bigrams_top.write(row, 0, chr(bigr[1]))
        symbol_bigrams_top.write(row, 1, chr(bigr[2]))
        symbol_bigrams_top.write(row, 2, bigr[3])
        symbol_bigrams_top.write(row, 3, bigr[3] / all_sum)
        if pos:
            cursor.execute(
                f'''
                SELECT AVG(position)
                FROM symbol_bigram_entries
                WHERE symbol_bigram_id = {str(bigr[0])};
                '''
            )
            symbol_bigrams_top.write(row, 4, cursor.fetchone()[0])


def sheet_ru_symb_bigrams(cursor):
    ru_symb_bigrams = workbook.add_worksheet('Russian letters bigrams')
    locations: dict = {
        **{x: n for n, x in enumerate(range(1040, 1046), 0)},
        **{x: n for n, x in enumerate(range(1072, 1078), 0)},
        **{x: n for n, x in enumerate(range(1046, 1072), 7)},
        **{x: n for n, x in enumerate(range(1078, 1104), 7)},
        **{1025: 6, 1105: 6},
    }

    cursor.execute(
        '''
        SELECT *
        FROM symbol_bigrams
        WHERE (first_symb BETWEEN 1040 AND 1105 OR first_symb = 1025)
            AND (second_symb BETWEEN 1040 AND 1105 OR second_symb = 1025);
        '''
    )

    values: list = [[]]

    for pair in cursor.fetchall():
        loc_1 = locations[pair[1]]
        loc_2 = locations[pair[2]]
        while len(values) <= loc_1:
            values.append([])
        while len(values[loc_1]) <= loc_2:
            values[loc_1].append(0)
        values[loc_1][loc_2] += pair[3]

    rus_let = 'абвгдеёжзийклмнопрстуфхцчшщьыъэюя'
    for n, let in enumerate(rus_let, 1):
        ru_symb_bigrams.write(0, n, let)
        ru_symb_bigrams.write(n, 0, let)

    for row, row_list in enumerate(values, 1):
        for col, col_value in enumerate(row_list, 1):
            ru_symb_bigrams.write(row, col, col_value)


def sheet_all_symb_bigrams(cursor, min_quantity):
    all_symb_bigrams = workbook.add_worksheet('All symb bigrams')
    cursor.execute(f'SELECT * FROM symbol_bigrams WHERE quantity >= {min_quantity}')
    positions: dict = {}
    for pair in cursor.fetchall():
        for elem_num in [1, 2]:
            if pair[elem_num] not in positions:
                positions[pair[elem_num]] = len(positions) + 1
                all_symb_bigrams.write(0, positions[pair[elem_num]], chr(pair[elem_num]))
                all_symb_bigrams.write(positions[pair[elem_num]], 0, chr(pair[elem_num]))
        all_symb_bigrams.write(positions[pair[1]], positions[pair[2]], pair[3])


def sheet_top_words(cursor, limit, min_quantity, pos=False):
    top_words = workbook.add_worksheet('Top words')
    top_words.write(0, 0, 'Word')
    top_words.write(0, 1, 'Quantity')
    top_words.write(0, 2, r'% from all')
    if pos:
        top_words.write(0, 3, 'As first')
        top_words.write(0, 4, 'Avg. position')

    cursor.execute('SELECT SUM(quantity) FROM words;')
    all_sum = cursor.fetchone()[0]

    cursor.execute(
        f'''
        SELECT *
        FROM words
        WHERE quantity >= {min_quantity}
        ORDER BY quantity DESC, word ASC
        {f'LIMIT {limit}' if limit else ''};
        '''
    )

    for row, word in enumerate(cursor.fetchall(), 1):
        top_words.write(row, 0, word[1])
        top_words.write(row, 1, word[2])
        top_words.write(row, 2, word[2] / all_sum)
        if pos:
            cursor.execute(
                f'''
                SELECT COUNT(*)
                FROM word_entries
                WHERE word_id = {str(word[0])} AND position = 1;
                '''
            )
            top_words.write(row, 3, cursor.fetchone()[0])
            cursor.execute(
                f'SELECT AVG(position) FROM word_entries WHERE word_id = {str(word[0])};'
            )
            top_words.write(row, 4, cursor.fetchone()[0])


def sheet_word_bigrams_top(cursor, limit, min_quantity, pos=False):
    word_bigrams_top = workbook.add_worksheet('Word bigrams top')
    word_bigrams_top.write(0, 0, 'First word')
    word_bigrams_top.write(0, 1, 'Second word')
    word_bigrams_top.write(0, 2, 'Quantity')
    word_bigrams_top.write(0, 3, r'% from all')
    if pos:
        word_bigrams_top.write(0, 4, 'Avg. position')

    cursor.execute('SELECT SUM(quantity) FROM word_bigrams;')
    all_sum = cursor.fetchone()[0]

    cursor.execute(
        f'''
        SELECT *
        FROM word_bigrams
        WHERE quantity >= {min_quantity}
        ORDER BY quantity DESC
        {f'LIMIT {limit}' if limit else ''};
        '''
    )

    for row, bigr in enumerate(cursor.fetchall(), 1):
        cursor.execute(f'SELECT word FROM words WHERE id={bigr[1]};')
        word_bigrams_top.write(row, 0, cursor.fetchone()[0])
        cursor.execute(f'SELECT word FROM words WHERE id={bigr[2]};')
        word_bigrams_top.write(row, 1, cursor.fetchone()[0])
        word_bigrams_top.write(row, 2, bigr[3])
        word_bigrams_top.write(row, 3, bigr[3] / all_sum)
        if pos:
            cursor.execute(
                f'''
                SELECT AVG(position)
                FROM word_bigram_entries
                WHERE word_bigram_id = {str(bigr[0])};
                '''
            )
            word_bigrams_top.write(row, 4, cursor.fetchone()[0])

def yo():
    pass



db = sqlite3.connect('app - Copy.db')
cur = db.cursor()

workbook = xlsxwriter.Workbook('result.xlsx')
sheet_symbols()
sheet_bigrams_top()
sheet_ru_bigrams()
sheet_all_bigrams()
sheet_words()
workbook.close()
