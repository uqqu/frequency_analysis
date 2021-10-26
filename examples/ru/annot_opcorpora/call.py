import io
from datetime import datetime
from os import listdir
from bs4 import BeautifulSoup

import frequency_analysis


start = datetime.now()
file_list = listdir('annot_opcorpora_xml_byfile/')

with frequency_analysis.Analysis(mode='n', yo=1) as analyze:
    for n, file in enumerate(file_list):
        with io.open('annot_opcorpora_xml_byfile/' + file, mode='r', encoding='utf-8') as f:
            data = f.read()
        bs_data = BeautifulSoup(data, 'xml')

        for sentence in bs_data.find_all('source'):
            analyze.count_all(sentence.text.split(), pos=True)
        print(n, file)
print('fin at:', datetime.now().strftime('%H:%M:%S'))
print('total time taked to analysis:', datetime.now() - start)
