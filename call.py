import io
from datetime import datetime
from os import listdir
from bs4 import BeautifulSoup

import frequency


file_list = listdir('multiUN/')
clear_word_pattern = '[^a-zA-Z’\'-]|^[^a-zA-Z]+|[^a-zA-Z]+$'
allowed_symbols = [*range(32, 127)]
intraword_symbols = {*range(65, 91), *range(97, 123)}

analyze = frequency.Analysis.open(
    mode='n',
    clear_word_pattern=clear_word_pattern,
    allowed_symbols=allowed_symbols,
    intraword_symbols=intraword_symbols,
)

for n, file in enumerate(file_list):
    with io.open('multiUN/' + file, mode='r', encoding='utf-8') as f:
        data = f.read()
    bs_data = BeautifulSoup(data, "xml")

    for sentence in bs_data.find_all("s"):
        analyze.count_all(sentence.text.split(), pos=True)
    print(n)
analyze.final()
print(datetime.now().strftime("%H:%M:%S"))
