import io
from os import listdir
from bs4 import BeautifulSoup

import frequency


file_list = sorted(listdir('annot.opcorpora.xml.byfile/'), key=lambda x: int(x.split('.')[0]))
clear_word_pattern = '[^а-яА-ЯёЁ’\'-]|^[^а-яА-ЯёЁ]+|[^а-яА-ЯёЁ]+$'
allowed_symbols = [*range(32, 65), 1025, *range(1040, 1104), 1105]
intraword_symbols = {45, 1025, *range(1040, 1104), 1105}

analyze = frequency.Analysis.open(
    mode='c',
    clear_word_pattern=clear_word_pattern,
    allowed_symbols=allowed_symbols,
    intraword_symbols=intraword_symbols,
    yo=True,
)

for file in file_list:
    with io.open('annot.opcorpora.xml.byfile/' + file, mode='r', encoding='utf-8') as f:
        data = f.read()
    bs_data = BeautifulSoup(data, "xml")

    for paragraph in bs_data.find_all("paragraph"):
        for sentence in paragraph.find_all("source"):
            analyze.count_all(sentence.text.split())
    print(file)
