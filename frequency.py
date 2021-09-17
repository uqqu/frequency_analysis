import os
import re
import sqlite3
from typing import Set, List, Union

import db_create


class FrequencyAnalysis:
    def __init__(
        self,
        name,
        mode,
        clear_word_pattern,
        intraword_symbols,
        total_symbols,
        total_words,
        db,
    ):
        self.name = name
        self.mode = mode
        self.clear_word_pattern = clear_word_pattern
        self.intraword_symbols = intraword_symbols
        self.total_symbols = total_symbols
        self.total_words = total_words
        self.db = db
        self.cursor = db.cursor()

    def count_all(self, word_list: list):
        self.count_words(word_list)
        for word in word_list:
            self.cursor.execute(
                '''
                UPDATE symbols
                SET quantity=quantity+1
                WHERE ord=32;
                '''
            )
            self.count_symbols(word)

    def count_words(self, word_list: list):
        last_word_id = None
        shift = 0
        for word_pos, word in enumerate(word_list):
            clear_word = re.sub(self.clear_word_pattern, '', word)
            clear_word = re.sub(self.clear_word_pattern, '', clear_word) # TODO
            if not clear_word:
                last_word_id = None
                shift += 1
                continue
            if self.total_words > 0:
                self.total_words -= 1
                continue

            self.cursor.execute(
                f'''
                SELECT id
                FROM words
                WHERE word="{clear_word.lower()}";
                '''
            )
            if res := self.cursor.fetchone():
                word_id = res[0]
                self.cursor.execute(
                    f'''
                    UPDATE words
                    SET quantity=quantity+1
                    WHERE id={word_id};
                    '''
                )
            else:
                self.cursor.execute(
                    f'''
                    INSERT INTO words (word, quantity)
                    VALUES ("{clear_word.lower()}", 1)
                    RETURNING id;
                    '''
                )
                word_id = self.cursor.fetchone()[0]

            self.cursor.execute(
                f'''
                INSERT INTO word_entries (word_id, position)
                VALUES ({word_id}, {word_pos - shift + 1});
                '''
            )
            if last_word_id:
                self.count_word_bigrams(last_word_id, word_id, word_pos)
            last_word_id = word_id
            self.db.commit()

    def count_word_bigrams(self, first_word_id: int, second_word_id: int, position: int):
        self.cursor.execute(
            f'''
            SELECT id
            FROM word_bigrams
            WHERE first_word_id="{first_word_id}"
                AND second_word_id="{second_word_id}";
            '''
        )
        if res := self.cursor.fetchone():
            word_bigram_id = res[0]
            self.cursor.execute(
                f'''
                UPDATE word_bigrams
                SET quantity=quantity+1
                WHERE id={word_bigram_id};
                '''
            )
        else:
            self.cursor.execute(
                f'''
                INSERT INTO word_bigrams (first_word_id, second_word_id, quantity)
                VALUES ("{first_word_id}", "{second_word_id}", 1)
                RETURNING id;
                '''
            )
            word_bigram_id = self.cursor.fetchone()[0]
        self.cursor.execute(
            f'''
            INSERT INTO word_bigram_entries (word_bigram_id, position)
            VALUES ({word_bigram_id}, {position});
            '''
        )

    def count_symbols(self, word: str):
        clear_word = re.sub(self.clear_word_pattern, '', word) # TODO
        if not (clear_word := re.sub(self.clear_word_pattern, '', clear_word)):
            return
        self.cursor.execute(
            f'''
            UPDATE symbols
            SET as_first=as_first+1
            WHERE ord={ord(clear_word[0])};
            '''
        )
        self.cursor.execute(
            f'''
            UPDATE symbols
            SET as_last=as_last+1
            WHERE ord={ord(clear_word[-1])};
            '''
        )
        last_symb_id = None
        shift = 0
        for symb_pos, symb in enumerate(word):
            if self.total_symbols > 0:
                self.total_symbols -= 1
                continue
            order = ord(symb)
            self.cursor.execute(
                f'''
                UPDATE symbols
                SET quantity=quantity+1
                WHERE ord={order}
                RETURNING id;
                '''
            )
            if not (res := self.cursor.fetchone()):
                last_symb_id = None
            else:
                symb_id = res[0]
                if order in self.intraword_symbols:
                    pos = symb_pos - shift + 1
                else:
                    shift += 1
                    pos = symb_pos + 1
                self.cursor.execute(
                    f'''
                    INSERT INTO symbol_entries (symb_id, position)
                    VALUES ({symb_id}, {pos});
                    '''
                )
                if last_symb_id:
                    self.count_symbol_bigrams(last_symb_id, symb_id, pos)
                last_symb_id = symb_id
        if last_symb_id:
            self.db.commit()

    def count_symbol_bigrams(self, first_symb_id: int, second_symb_id: int, position: int):
        self.cursor.execute(
            f'''
            SELECT id
            FROM symbol_bigrams
            WHERE first_symb_id="{first_symb_id}"
                AND second_symb_id="{second_symb_id}";
            '''
        )
        if res := self.cursor.fetchone():
            symbol_bigram_id = res[0]
            self.cursor.execute(
                f'''
                UPDATE symbol_bigrams
                SET quantity=quantity+1
                WHERE id={symbol_bigram_id};
                '''
            )
        else:
            self.cursor.execute(
                f'''
                INSERT INTO symbol_bigrams (first_symb_id, second_symb_id, quantity)
                VALUES ("{first_symb_id}", "{second_symb_id}", 1)
                RETURNING id;
                '''
            )
            symbol_bigram_id = self.cursor.fetchone()[0]
        self.cursor.execute(
            f'''
            INSERT INTO symbol_bigram_entries (symbol_bigram_id, position)
            VALUES ({symbol_bigram_id}, {position});
            '''
        )


class Analysis:
    @staticmethod
    def open(
        name: str = 'frequency_analysis',
        mode: str = 'n',  # n – new file, a – append to existing, c – continue to existing
        clear_word_pattern: str = '[^а-яА-ЯёЁa-zA-Z’\'-]|^[\'-]|[\'-]$',
        allowed_symbols: List[Union[int, str]] = [*range(32, 127), 1025, *range(1040, 1104), 1105],
        intraword_symbols: Set[Union[int, str]] = {
            45,
            *range(64, 91),
            *range(97, 123),
            1025,
            *range(1040, 1104),
            1105,
        },
        yo: bool = False,
    ):
        if not re.search('^[a-zа-яё0-9_.@() -]+$', name, re.I):
            raise Exception(f"Filename '{name}' is unvalid. Please, enter other.")
        if mode not in ('n', 'a', 'c'):
            raise Exception(
                "Mode must be 'n' for new analysis, 'a' for append to existing \
                    or 'c' to continue the previous analysis. If empty – works as 'n'."
            )
        if mode == 'n' and os.path.isfile(f'./{name}.db'):
            raise Exception(
                f"DB file with the '{name}' name already exist! Use mode 'a' \
                    for append to existing, or 'c' to continue the previous analysis."
            )
        if mode != 'n' and not os.path.isfile(f'./{name}.db'):
            raise Exception(
                "Pattern for cleaning words is broken. \
                    Use mode 'n' (or leave it empty) to create a new analysis, \
                    or set name of existing DB (without extension)."
            )
        try:
            re.compile(clear_word_pattern)
        except re.error as re_error:
            raise Exception(
                f"DB file with the '{name}' name is not exist! \
                    Use mode 'n' to create a new analysis, or set name of existing DB \
                    (without extension)."
            ) from re_error
        if (
            not isinstance(allowed_symbols, (str, list))
            or isinstance(allowed_symbols, list)
            and not (
                all(isinstance(x, int) for x in allowed_symbols)
                or all(isinstance(x, str) for x in allowed_symbols)
            )
        ):
            raise Exception(
                "Allowed symbols must be a string or a list of single symbols \
                    or a list of integers (decimal unicode values). \
                    If empty works as <base latin> + <russian cyrillic> + <numbers> + \
                    <space> + '!\"#$%&'()*+,-./:;<>=?@[]\\^_`{}|~'."
            )
        if (
            not isinstance(intraword_symbols, (str, set))
            or isinstance(intraword_symbols, set)
            and not (
                all(isinstance(x, int) for x in intraword_symbols)
                or all(isinstance(x, str) for x in intraword_symbols)
            )
        ):
            raise Exception(
                "Intraword symbols must be a string or a set of single symbols \
                    or a set of integers (decimal unicode values). \
                    If empty works as <base latin> + <russian cyrillic> + '-'."
            )
        if yo and (not os.path.isfile('./yo.txt') or not os.path.isfile('./ye-yo.txt')):
            raise Exception(
                "Yo mode require additional 'yo.txt' and 'ye-yo.txt' files near the script."
            )

        if isinstance(allowed_symbols[0], str):
            allowed_symbols = [ord(x) for x in allowed_symbols]
        if isinstance(intraword_symbols, str) or isinstance(next(iter(intraword_symbols)), str):
            intraword_symbols = {ord(x) for x in intraword_symbols}

        total_words = 0
        total_symbols = 0
        db = sqlite3.connect(f'{name}.db')
        cursor = db.cursor()
        if mode == 'n':
            db_create.create_new(db, allowed_symbols)
            if yo:
                db_create.yo_mode(db)
        elif mode == 'c':
            cursor.execute('SELECT SUM(quantity) FROM words;')
            total_words = cursor.fetchone()[0]
            cursor.execute('SELECT SUM(quantity) FROM symbols;')
            total_symbols = cursor.fetchone()[0]

        return FrequencyAnalysis(
            name, mode, clear_word_pattern, intraword_symbols, total_symbols, total_words, db
        )


__all__ = ['Analysis']
