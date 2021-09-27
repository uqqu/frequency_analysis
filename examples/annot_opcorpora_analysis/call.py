import io
from datetime import datetime
from os import listdir
from bs4 import BeautifulSoup

import frequency_analysis


file_list = listdir('annot_opcorpora_xml_byfile/')

with frequency_analysis.Analysis(mode='n', yo=True) as analyze:
    for n, file in enumerate(file_list):
        with io.open('annot_opcorpora_xml_byfile/' + file, mode='r', encoding='utf-8') as f:
            data = f.read()
        bs_data = BeautifulSoup(data, 'xml')

        for sentence in bs_data.find_all('source'):
            analyze.count_all(sentence.text.split(), pos=True)
        print(n, file)
print(datetime.now().strftime('%H:%M:%S'))

with frequency_analysis.Result() as res:
    res.treat(limits=[1000]*4, min_quantity=[10]*5)
    res.sheet_custom_symb(''.join([chr(x) for x in range(1072, 1104)] + [chr(1105)]))
    res.sheet_ru_symb_bigrams()
    res.sheet_yo_words()
print(datetime.now().strftime('%H:%M:%S'))
