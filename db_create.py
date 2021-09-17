import io


def create_new(db, allowed_symbols):
    cursor = db.cursor()
    cursor.execute(
        '''
        CREATE TABLE symbols (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ord INTEGER NOT NULL UNIQUE,
        quantity INTEGER NOT NULL,
        as_first INTEGER NOT NULL,
        as_last INTEGER NOT NULL);
        '''
    )
    cursor.execute(
        '''
        CREATE TABLE symbol_entries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            symb_id INTEGER NOT NULL,
            position INTEGER NOT NULL,
            FOREIGN KEY (symb_id)
                REFERENCES symbols (id));
        '''
    )
    cursor.execute(
        '''
        CREATE TABLE symbol_bigrams (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            first_symb_id INTEGER NOT NULL,
            second_symb_id INTEGER NOT NULL,
            quantity INTEGER NOT NULL,
            FOREIGN KEY (first_symb_id)
                REFERENCES symbols (id),
            FOREIGN KEY (second_symb_id)
                REFERENCES symbols (id));
        '''
    )
    cursor.execute(
        '''
        CREATE TABLE symbol_bigram_entries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            symbol_bigram_id INTEGER NOT NULL,
            position INTEGER NOT NULL,
            FOREIGN KEY (symbol_bigram_id)
                REFERENCES symbol_bigrams (id));
        '''
    )
    cursor.execute(
        '''
        CREATE TABLE words (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            word TEXT NOT NULL,
            quantity INTEGER NOT NULL);
        '''
    )
    cursor.execute(
        '''
        CREATE TABLE word_entries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            word_id INTEGER NOT NULL,
            position INTEGER NOT NULL,
            FOREIGN KEY (word_id)
                REFERENCES words (id));
        '''
    )
    cursor.execute(
        '''
        CREATE TABLE word_bigrams (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            first_word_id INTEGER NOT NULL,
            second_word_id INTEGER NOT NULL,
            quantity INTEGER NOT NULL,
            FOREIGN KEY (first_word_id)
                REFERENCES words (id),
            FOREIGN KEY (second_word_id)
                REFERENCES words (id));
        '''
    )
    cursor.execute(
        '''
        CREATE TABLE word_bigram_entries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            word_bigram_id INTEGER NOT NULL,
            position INTEGER NOT NULL,
            FOREIGN KEY (word_bigram_id)
                REFERENCES word_bigrams (id));
        '''
    )

    db.commit()

    for symb_ord in allowed_symbols:
        cursor.execute(
            f'''
            INSERT INTO symbols (ord, quantity, as_first, as_last)
            VALUES ("{symb_ord}", 0, 0, 0);
            '''
        )
    db.commit()


def yo_mode(db):
    cursor = db.cursor()
    cursor.execute(
        '''
        CREATE TABLE yo_words (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            yo_word_id INTEGER NOT NULL,
            ye_word_id INTEGER NOT NULL,
            mandatory BOOLEAN,
            FOREIGN KEY (yo_word_id)
                REFERENCES words (id),
            FOREIGN KEY (ye_word_id)
                REFERENCES words (id));
        '''
    )

    with io.open('yo.txt', mode='r', encoding='utf-8') as f:
        for line in f:
            cursor.execute(
                f'''
                INSERT INTO words (word, quantity)
                VALUES ("{line.strip()}", 0)
                RETURNING id;
                '''
            )
            yo_word_id = cursor.fetchone()[0]
            value = line.strip().replace('ё', 'е')
            cursor.execute(
                f'''
                INSERT INTO words (word, quantity)
                VALUES ("{value}", 0)
                RETURNING id;
                '''
            )
            ye_word_id = cursor.fetchone()[0]

            cursor.execute(
                f'''
                INSERT INTO yo_words (yo_word_id, ye_word_id, mandatory)
                VALUES ("{yo_word_id}", "{ye_word_id}", 1)
                RETURNING id;
                '''
            )

    with io.open('ye-yo.txt', mode='r', encoding='utf-8') as f:
        for line in f:
            cursor.execute(
                f'''
                INSERT INTO words (word, quantity)
                VALUES ("{line.strip()}", 0)
                RETURNING id;
                '''
            )
            yo_word_id = cursor.fetchone()[0]
            value = line.strip().replace('ё', 'е')
            cursor.execute(
                f'''
                INSERT INTO words (word, quantity)
                VALUES ("{value}", 0)
                RETURNING id;
                '''
            )
            ye_word_id = cursor.fetchone()[0]

            cursor.execute(
                f'''
                INSERT INTO yo_words (yo_word_id, ye_word_id, mandatory)
                VALUES ("{yo_word_id}", "{ye_word_id}", 0)
                RETURNING id;
                '''
            )
