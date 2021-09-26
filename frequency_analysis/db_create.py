'''Additional module to frequency.py for creating separate DB for each analysis.'''

import io


def create_new(db, allowed_symbols):
    '''Create all necessary tables.'''
    cursor = db.cursor()
    cursor.execute(
        '''
        CREATE TABLE IF NOT EXISTS symbols (
            ord INTEGER PRIMARY KEY,
            quantity INTEGER NOT NULL,
            as_first INTEGER NOT NULL,
            as_last INTEGER NOT NULL,
            position REAL
        ) WITHOUT ROWID;
        '''
    )
    cursor.execute(
        '''
        CREATE TABLE symbol_bigrams (
            first_symb_ord INTEGER,
            second_symb_ord INTEGER,
            quantity INTEGER NOT NULL,
            position REAL,
            PRIMARY KEY (first_symb_ord, second_symb_ord),
            FOREIGN KEY (first_symb_ord)
                REFERENCES symbols (ord),
            FOREIGN KEY (second_symb_ord)
                REFERENCES symbols (ord)
        ) WITHOUT ROWID;
        '''
    )
    cursor.execute(
        '''
        CREATE TABLE words (
            word TEXT PRIMARY KEY,
            quantity INTEGER NOT NULL,
            as_first INTEGER NOT NULL,
            as_last INTEGER NOT NULL,
            position REAL
        ) WITHOUT ROWID;
        '''
    )
    cursor.execute(
        '''
        CREATE TABLE word_bigrams (
            first_word TEXT,
            second_word TEXT,
            quantity INTEGER NOT NULL,
            position REAL,
            PRIMARY KEY (first_word, second_word),
            FOREIGN KEY (first_word)
                REFERENCES words (word),
            FOREIGN KEY (second_word)
                REFERENCES words (word)
        ) WITHOUT ROWID;
        '''
    )

    db.commit()

    for symb_ord in allowed_symbols:
        cursor.execute(
            f'''
            INSERT INTO symbols (ord, quantity, as_first, as_last, position)
            VALUES ("{symb_ord}", 0, 0, 0, 1);
            '''
        )
    db.commit()


def yo_mode(db):
    '''Create additional table for a demonstration ye/yo Cyrillic misspelling.

    Require additional files with ye/yo word lists.
    One of the options: https://github.com/uqqu/yo_dict'''
    cursor = db.cursor()
    cursor.execute(
        '''
        CREATE TABLE yo_words (
            yo_word TEXT,
            ye_word TEXT,
            mandatory BOOLEAN,
            PRIMARY KEY (yo_word, ye_word),
            FOREIGN KEY (yo_word)
                REFERENCES words (word),
            FOREIGN KEY (ye_word)
                REFERENCES words (word)
        ) WITHOUT ROWID;
        '''
    )

    with io.open('yo.txt', mode='r', encoding='utf-8') as f:
        for line in f:
            yo_word = line.strip()
            cursor.execute(
                f'''
                INSERT INTO words (word, quantity, as_first, as_last, position)
                VALUES ("{yo_word}", 0, 0, 0, 1)
                ON CONFLICT DO NOTHING;
                '''
            )
            ye_word = yo_word.replace('ё', 'е')
            cursor.execute(
                f'''
                INSERT INTO words (word, quantity, as_first, as_last, position)
                VALUES ("{ye_word}", 0, 0, 0, 1)
                ON CONFLICT DO NOTHING;
                '''
            )

            cursor.execute(
                f'''
                INSERT INTO yo_words (yo_word, ye_word, mandatory)
                VALUES ("{yo_word}", "{ye_word}", 1)
                ON CONFLICT DO NOTHING;
                '''
            )

    with io.open('ye-yo.txt', mode='r', encoding='utf-8') as f:
        for line in f:
            yo_word = line.strip()
            cursor.execute(
                f'''
                INSERT INTO words (word, quantity, as_first, as_last, position)
                VALUES ("{yo_word}", 0, 0, 0, 1)
                ON CONFLICT DO NOTHING;
                '''
            )
            ye_word = line.strip().replace('ё', 'е')
            cursor.execute(
                f'''
                INSERT INTO words (word, quantity, as_first, as_last, position)
                VALUES ("{yo_word}", 0, 0, 0, 1)
                ON CONFLICT DO NOTHING;
                '''
            )

            cursor.execute(
                f'''
                INSERT INTO yo_words (yo_word, ye_word, mandatory)
                VALUES ("{yo_word}", "{ye_word}", 0)
                ON CONFLICT DO NOTHING;
                '''
            )
