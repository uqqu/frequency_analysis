﻿'''Main module for frequency analysis.'''

import os
import re
import sqlite3
from typing import List, Union

from frequency_analysis import db_create


def commit(func):
    '''Decorator for commit changes once per 100 cycles.'''

    def inner(self, *args, **kwargs):
        self.counter += 1
        if self.counter > 100:
            self.counter = 0
            self.db.commit()
        return func(self, *args, **kwargs)

    return inner


class FrequencyAnalysis:
    '''End-user class to perform frequency analysis for user data/corpus.

    User input:
        /all values are optional/
        name            – the name for the analysis folder;
        word_pattern    – regex pattern to extract words from a sentence;
        allowed_symbols – symbols which will be taken into account in the process of analysis.
    '''

    def __init__(self, name, word_pattern, allowed_symbols, total_symbols, total_words, db):
        self.name = name
        self.word_pattern = word_pattern
        self.allowed_symbols = allowed_symbols
        self.total_symbols = total_symbols
        self.total_words = total_words
        self.db = db
        self.cursor = db.cursor()
        self.counter = 0
        self.space = ' ' in self.allowed_symbols

    def __create_clear_word_list(self, word_list):
        shift = 0
        clear_word_list = []
        for n, word in enumerate(word_list[:], 1):
            if res := re.findall(self.word_pattern, word):
                [clear_word_list.append(x) for x in res]
                [word_list.insert(n + shift, '') for _ in range(len(res) - 1)]
                shift += len(res) - 1
            else:
                clear_word_list.append('')
        return clear_word_list

    @commit
    def count_all(self, word_list: list, pos=False, symbol_bigrams=True, word_bigrams=True):
        '''Count symbols, words, symbol bigrams, word bigrams, all their average positions.

        Input:
            Word list – sentence for analysis.
                It must be exactly sentence for properly word position counting;
            Average position counting – disabled by default. Slows down performance by ≈20%;
            Symbol bigrams counting – enabled by default;
            Word bigrams counting – enabled by default.
        '''
        clear_word_list = self.__create_clear_word_list(word_list)
        for word, clear_word in zip(word_list, clear_word_list):
            if self.total_symbols == 0 and self.space:
                self.cursor.execute(
                    '''
                    UPDATE symbols
                    SET quantity=quantity+1
                    WHERE chr=' ';
                    '''
                )
            self.__count_symbols(word, clear_word, pos, symbol_bigrams)

        if cutted_clear_word_list := [x.replace("'", "''") for x in clear_word_list if x]:
            self.__count_words(cutted_clear_word_list, pos, word_bigrams)

    @commit
    def count_words(self, word_list: list, pos=False, bigrams=True):
        '''Decorated wrapper for user calling.'''
        clear_word_list = self.__create_clear_word_list(word_list)
        if cutted_clear_word_list := [x.replace("'", "''") for x in clear_word_list if x]:
            self.__count_words(cutted_clear_word_list, pos, bigrams)

    @commit
    def count_symbols(self, word_list: list, pos=False, bigrams=True):
        '''Decorated wrapper for user calling.'''
        for word, clear_word in zip(word_list, self.__create_clear_word_list(word_list)):
            if self.total_symbols == 0 and self.space:
                self.cursor.execute(
                    '''
                    UPDATE symbols
                    SET quantity=quantity+1
                    WHERE chr=' ';
                    '''
                )
            self.__count_symbols(word, clear_word, pos, bigrams)

    def __count_words(self, word_list: list, pos: bool, bigrams: bool):
        '''Word/word bigrams counting.'''
        last_word = None
        for word_pos, word in enumerate(word_list, 1):
            if self.total_words > 0:
                self.total_words -= 1
                continue
            self.cursor.execute(
                f'''
                INSERT INTO words
                    (word, quantity, as_first, as_last {', position' if pos else ''})
                VALUES ('{word.lower()}', 1, 0, 0 {f', {word_pos}' if pos else ''})
                ON CONFLICT (word) DO UPDATE SET quantity=quantity+1
                    {f', position=(position*quantity+{word_pos}) / (quantity+1)' if pos else ''};
                '''
            )
            if last_word and bigrams:
                self.cursor.execute(
                    f'''
                    INSERT INTO word_bigrams
                        (first_word, second_word, quantity, as_first, as_last
                            {', position' if pos else ''})
                    VALUES ('{last_word}', '{word.lower()}', 1, 0, 0
                        {f', {word_pos-1}' if pos else ''})
                    ON CONFLICT (first_word, second_word) DO UPDATE SET quantity=quantity+1
                    {f', position=(position*quantity+{word_pos-1}) / (quantity+1)' if pos else ''};
                    '''
                )
            last_word = word.lower()
        if len(word_list) > 1 and not self.total_words:
            self.cursor.execute(
                f'''
                UPDATE words
                SET as_first=as_first+1
                WHERE word='{word_list[0].lower()}';
                '''
            )
            self.cursor.execute(
                f'''
                UPDATE words
                SET as_last=as_last+1
                WHERE word='{word_list[-1].lower()}';
                '''
            )
            if len(word_list) > 2 and bigrams:
                self.cursor.execute(
                    f'''
                    UPDATE word_bigrams
                    SET as_first=as_first+1
                    WHERE first_word='{word_list[0].lower()}'
                        AND second_word='{word_list[1].lower()}';
                    '''
                )
                self.cursor.execute(
                    f'''
                    UPDATE word_bigrams
                    SET as_last=as_last+1
                    WHERE first_word='{word_list[-2].lower()}'
                        AND second_word='{word_list[-1].lower()}';
                    '''
                )

    def __count_symbols(self, word: str, clear_word: str, pos: bool, bigrams: bool):
        '''Symbol/symbol bigrams counting.'''
        last_symb = None
        shift = 0
        for symb_pos, symb in enumerate(word, 1):
            if symb not in self.allowed_symbols:
                last_symb = None
                continue
            if self.total_symbols > 0:
                self.total_symbols -= 1
                continue
            if re.search(self.word_pattern, symb):
                position = symb_pos - shift
            else:
                shift += 1
                position = symb_pos
            symb = symb.replace("'", "''")
            self.cursor.execute(
                f'''
                UPDATE symbols
                SET quantity=quantity+1
                    {f', position=(position*quantity+{position}) / (quantity+1)' if pos else ''}
                WHERE chr='{symb}';
                '''
            )

            if last_symb and bigrams:
                self.cursor.execute(
                    f'''
                    INSERT INTO symbol_bigrams
                        (first_symb, second_symb, quantity, as_first, as_last
                            {', position' if pos else ''})
                    VALUES ('{last_symb}', '{symb}', 1, 0, 0 {f', {position - 1}' if pos else ''})
                    ON CONFLICT (first_symb, second_symb)
                        DO UPDATE SET quantity=quantity+1
                    {f', position=(position*quantity+{position-1}) / (quantity+1)' if pos else ''};
                    '''
                )

            last_symb = symb

        if len(clear_word) > 1 and not self.total_symbols:
            self.cursor.execute(
                f'''
                UPDATE symbols
                SET as_first=as_first+1
                WHERE chr='{clear_word[0]}';
                '''
            )
            self.cursor.execute(
                f'''
                UPDATE symbols
                SET as_last=as_last+1
                WHERE chr='{clear_word[-1]}';
                '''
            )
            if len(clear_word) > 2 and bigrams:
                values = [x.replace("'", "''") for x in clear_word]
                if values[0] in self.allowed_symbols and values[1] in self.allowed_symbols:
                    self.cursor.execute(
                        f'''
                        UPDATE symbol_bigrams
                        SET as_first=as_first+1
                        WHERE first_symb='{values[0]}' AND second_symb='{values[1]}';
                        '''
                    )
                if values[-1] in self.allowed_symbols and values[-2] in self.allowed_symbols:
                    self.cursor.execute(
                        f'''
                        UPDATE symbol_bigrams
                        SET as_last=as_last+1
                        WHERE first_symb='{values[-2]}' AND second_symb='{values[-1]}'
                        '''
                    )


class Analysis:
    '''Context manager with data validation for end-user FrequencyAnalysis class.'''

    def __init__(
        self,
        name: str = 'frequency_analysis',
        mode: str = 'n',  # n – new file, a – append to existing, c – continue to existing
        word_pattern: str = '[a-zA-Zа-яА-ЯёЁ]+(?:(?:-?[a-zA-Zа-яА-ЯёЁ]+)+|\
                \'?[a-zA-Zа-яА-ЯёЁ]+)|[a-zA-Zа-яА-ЯёЁ]',
        allowed_symbols: List[Union[int, str]] = [*range(32, 127), 1025, *range(1040, 1104), 1105],
        yo: int = 0,
    ):
        self.name = name
        self.mode = mode
        self.word_pattern = word_pattern
        self.allowed_symbols = allowed_symbols
        self.yo = yo
        self.db = None

    def __enter__(self):
        if not re.search('^[a-zа-яё0-9_.@() -]+$', self.name, re.I):
            raise Exception(f"Foldername '{self.name}' is unvalid. Please, enter other.")
        if self.mode not in ('n', 'a', 'c'):
            raise Exception(
                "Mode must be 'n' for new analysis, 'a' for append to existing "
                "or 'c' to continue the previous analysis. If empty – works as 'n'."
            )
        if self.mode == 'n' and os.path.isfile(os.path.join(os.getcwd(), self.name, 'result.db')):
            raise Exception(
                f"DB file in the '{self.name}' folder already exist! Use mode 'a' "
                "for append to existing, or 'c' to continue the previous analysis."
            )
        if self.mode != 'n' and not os.path.isfile(
            os.path.join(os.getcwd(), self.name, 'result.db')
        ):
            raise Exception(
                f"DB file in the '{self.name}' folder is not exist! "
                "Use mode 'n' to create a new analysis, or set name of folder with existing DB."
            )
        try:
            re.compile(self.word_pattern)
        except re.error as re_error:
            raise Exception(
                "Pattern for extracting words from a sentence is broken."
            ) from re_error
        if (
            not isinstance(self.allowed_symbols, (str, list))
            or isinstance(self.allowed_symbols, list)
            and not (
                all(isinstance(x, int) for x in self.allowed_symbols)
                or all(isinstance(x, str) for x in self.allowed_symbols)
            )
        ):
            raise Exception(
                "Allowed symbols must be a string or a list of single symbols "
                "or a list of integers (decimal unicode values). "
                "If empty works as <base latin> + <russian cyrillic> + <numbers> + "
                "<space> + '!\"#$%&'()*+,-./:;<>=?@[]\\^_`{}|~'."
            )
        if self.yo and (
            not os.path.isfile(os.path.join(os.getcwd(), 'yo.txt'))
            or not os.path.isfile(os.path.join(os.getcwd(), 'ye-yo.txt'))
        ):
            raise Exception(
                "Yo mode require additional 'yo.txt' and 'ye-yo.txt' files near the script."
            )

        if isinstance(self.allowed_symbols[0], int):
            self.allowed_symbols = [chr(x) for x in self.allowed_symbols]

        if not os.path.exists(os.path.join(os.getcwd(), self.name)):
            os.mkdir(os.path.join(os.getcwd(), self.name))

        total_words = 0
        total_symbols = 0
        self.db = sqlite3.connect(os.path.join(os.getcwd(), self.name, 'result.db'))
        cursor = self.db.cursor()
        if self.mode == 'n':
            db_create.create_new(self.db, self.allowed_symbols)
            if self.yo:
                db_create.yo_mode(self.db)
        elif self.mode == 'c':
            total_words = cursor.execute('SELECT SUM(quantity) FROM words;').fetchone()[0]
            total_symbols = cursor.execute('SELECT SUM(quantity) FROM symbols;').fetchone()[0]
        elif self.mode == 'a' and self.yo == 2:
            db_create.yo_mode(self.db, True)

        return FrequencyAnalysis(
            self.name, self.word_pattern, self.allowed_symbols, total_symbols, total_words, self.db
        )

    def __exit__(self, type_, value, traceback):
        self.db.commit()
        self.db.close()


__all__ = ['Analysis']
