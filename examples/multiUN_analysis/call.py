import io
from datetime import datetime
from os import listdir
from bs4 import BeautifulSoup

import frequency_analysis


start = datetime.now()
file_list = listdir('multiUN/')
word_pattern = '[a-zA-Z]+(?:(?:-?[a-zA-Z]+)+|\'?[a-zA-Z]+)|[a-zA-Z]'
allowed_symbols = [*range(32, 127)]

with frequency_analysis.Analysis(
    mode='c', word_pattern=word_pattern, allowed_symbols=allowed_symbols
) as analyze:
    for n, file in enumerate(file_list):
        with io.open('multiUN/' + file, mode='r', encoding='utf-8') as f:
            data = f.read()
        bs_data = BeautifulSoup(data, 'xml')

        for sentence in bs_data.find_all('s'):
            analyze.count_all(sentence.text.split(), pos=True)
        print(n, file)
print('fin at:', datetime.now().strftime('%H:%M:%S'))
print('total time:', datetime.now() - start)