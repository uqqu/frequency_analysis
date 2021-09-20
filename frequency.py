import os
import re
import sqlite3
from typing import Set, List, Union

import db_create


class FrequencyAnalysis:
    def __init__(
        self,
        name,
        clear_word_pattern,
        allowed_symbols,
        intraword_symbols,
        total_symbols,
        total_words,
        db,
    ):
        self.__name = name
        self.__clear_word_pattern = clear_word_pattern
        self.__allowed_symbols = allowed_symbols
        self.__intraword_symbols = intraword_symbols
        self.__total_symbols = total_symbols
        self.__total_words = total_words
        self.__db = db
        self.__cursor = db.cursor()
        self.__counter = 0

    def commit(func):
        def inner(self, *args, **kwargs):
            self.__counter += 1
            if self.__counter > 100:
                self.__counter = 0
                self.__db.commit()
            return func(self, *args, **kwargs)

        return inner

    def final(self):
        self.__db.commit()

    @commit
    def count_all(self, word_list: list, pos=False, symbol_bigram=True, word_bigram=True):
        clear_word_list = [re.sub(self.__clear_word_pattern, '', x).lower() for x in word_list]
        for word, clear_word in zip(word_list, clear_word_list):
            if self.__total_symbols == 0:
                self.__cursor.execute(
                    '''
                    UPDATE symbols
                    SET quantity=quantity+1
                    WHERE ord=32;
                    '''
                )
            self.__count_symbols(word, clear_word, pos, symbol_bigram)

        if cutted_clear_word_list := [x for x in clear_word_list if x]:
            self.__count_words(cutted_clear_word_list, pos, word_bigram)

    @commit
    def count_words(self, word_list: list, pos=False, bigram=True):
        clear_word_list = [re.sub(self.__clear_word_pattern, '', x).lower() for x in word_list]
        if cutted_clear_word_list := [x for x in clear_word_list if x]:
            self.__count_words(cutted_clear_word_list, pos, bigram)

    @commit
    def count_symbols(self, word_list: list, pos=False, bigram=True):
        clear_word_list = [re.sub(self.__clear_word_pattern, '', x).lower() for x in word_list]
        cutted_clear_word_list = [x for x in clear_word_list if x]

        for word, clear_word in zip(word_list, clear_word_list):
            if self.__total_symbols == 0:
                self.__cursor.execute(
                    '''
                    UPDATE symbols
                    SET quantity=quantity+1
                    WHERE ord=32;
                    '''
                )
            self.__count_symbols(word, clear_word, pos, bigram)

    def __count_words(self, word_list: list, pos: bool, bigram: bool):
        last_word = None
        for word_pos, word in enumerate(word_list, 1):
            if self.__total_words > 0:
                self.__total_words -= 1
                continue
            self.__cursor.execute(
                f'''
                INSERT INTO words
                    (word, quantity, as_first, as_last {', position' if pos else ''})
                VALUES ("{word}", 1, 0, 0 {f', {word_pos}' if pos else ''})
                ON CONFLICT (word) DO UPDATE SET quantity=quantity+1
                    {f', position=(position*quantity+{word_pos}) / (quantity+1)' if pos else ''};
                '''
            )
            if last_word and bigram:
                self.__cursor.execute(
                    f'''
                    INSERT INTO word_bigrams
                        (first_word, second_word, quantity {', position' if pos else ''})
                    VALUES ("{last_word}", "{word}", 1 {f', {word_pos - 1}' if pos else ''})
                    ON CONFLICT (first_word, second_word) DO UPDATE SET quantity=quantity+1
                    {f', position=(position*quantity+{word_pos - 1}) / (quantity+1)' if pos else ''};
                    '''
                )
            last_word = word
        self.__cursor.execute(
            f'''
            UPDATE words
            SET as_first=as_first+1
            WHERE word="{word_list[0]}";
            '''
        )
        self.__cursor.execute(
            f'''
            UPDATE words
            SET as_last=as_last+1
            WHERE word="{word_list[-1]}";
            '''
        )

    def __count_symbols(self, word: str, clear_word: str, pos: bool, bigram: bool):
        last_symb_ord = None
        shift = 0
        for symb_pos, symb in enumerate(word, 1):
            symb_ord = ord(symb)
            if symb_ord not in self.__allowed_symbols:
                last_symb_ord = None
                continue
            if self.__total_symbols > 0:
                self.__total_symbols -= 1
                continue
            if symb_ord in self.__intraword_symbols:
                position = symb_pos - shift
            else:
                shift += 1
                position = symb_pos
            self.__cursor.execute(
                f'''
                UPDATE symbols
                SET quantity=quantity+1
                {f', position=(position*quantity+{position}) / (quantity+1)' if pos else ''}
                WHERE ord={symb_ord};
                '''
            )

            if last_symb_ord and bigram:
                self.__cursor.execute(
                    f'''
                    INSERT INTO symbol_bigrams
                        (first_symb_ord, second_symb_ord, quantity {', position' if pos else ''})
                    VALUES ({last_symb_ord}, {symb_ord}, 1 {f', {position - 1}' if pos else ''})
                    ON CONFLICT (first_symb_ord, second_symb_ord) DO UPDATE SET quantity=quantity+1
                    {f', position=(position*quantity+{position - 1}) / (quantity+1)' if pos else ''};
                    '''
                )

            last_symb_ord = symb_ord

        if clear_word:
            self.__cursor.execute(
                f'''
                UPDATE symbols
                SET as_first=as_first+1
                WHERE ord={ord(clear_word[0])};
                '''
            )
            self.__cursor.execute(
                f'''
                UPDATE symbols
                SET as_last=as_last+1
                WHERE ord={ord(clear_word[-1])};
                '''
            )


class Analysis:
    @staticmethod
    def open(
        name: str = 'frequency_analysis',
        mode: str = 'n',  # n – new file, a – append to existing, c – continue to existing
        clear_word_pattern: str = '[^а-яА-ЯёЁa-zA-Z’\'-]|^[^а-яА-ЯёЁa-zA-Z]|[^а-яА-ЯёЁa-zA-Z]$',
        allowed_symbols: List[Union[int, str]] = [*range(32, 127), 1025, *range(1040, 1104), 1105],
        intraword_symbols: Set[Union[int, str]] = {
            45,
            *range(65, 91),
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
                "Mode must be 'n' for new analysis, 'a' for append to existing "
                "or 'c' to continue the previous analysis. If empty – works as 'n'."
            )
        if mode == 'n' and os.path.isfile(f'./{name}.db'):
            raise Exception(
                f"DB file with the '{name}' name already exist! Use mode 'a' "
                "for append to existing, or 'c' to continue the previous analysis."
            )
        if mode != 'n' and not os.path.isfile(f'./{name}.db'):
            raise Exception(
                "Pattern for cleaning words is broken. "
                "Use mode 'n' (or leave it empty) to create a new analysis, "
                "or set name of existing DB (without extension)."
            )
        try:
            re.compile(clear_word_pattern)
        except re.error as re_error:
            raise Exception(
                f"DB file with the '{name}' name is not exist! "
                "Use mode 'n' to create a new analysis, or set name of existing DB "
                "(without extension)."
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
                "Allowed symbols must be a string or a list of single symbols "
                "or a list of integers (decimal unicode values). "
                "If empty works as <base latin> + <russian cyrillic> + <numbers> + "
                "<space> + '!\"#$%&'()*+,-./:;<>=?@[]\\^_`{}|~'."
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
                "Intraword symbols must be a string or a set of single symbols "
                "or a set of integers (decimal unicode values). "
                "If empty works as <base latin> + <russian cyrillic> + '-'."
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
            name,
            clear_word_pattern,
            allowed_symbols,
            intraword_symbols,
            total_symbols,
            total_words,
            db,
        )


__all__ = ['Analysis']
