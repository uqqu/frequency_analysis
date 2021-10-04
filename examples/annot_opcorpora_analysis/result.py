from datetime import datetime
import frequency_analysis

start = datetime.now()
ru_symbs = 'абвгдеёжзийклмнопрстуфхцчшщьыъэюя'

with frequency_analysis.Result() as res:
    # treat() calls all main sheet functions.
    #   limits – max number of items for sheets (symbols, symbol bigrams, words, word bigrams)
    #   min_quantity – min quantity of each item for sheets
    #       (symbols, symbol bigrams, words, word bigrams, symbol bigrams 2D)
    #   chart_limit – number of first n items for charts on
    #       (symbols, symbol bigrams, words, word bigrams) sheets
    res.treat(limits=(1000,) * 4, min_quantity=(10,) * 5, chart_limit=(20,) * 4)
    # sheet_custom_symb(symbols: string/list/tuple/set, chart_limit: int, name: str (sheet_name)
    res.sheet_custom_symb(''.join([chr(x) for x in range(1040, 1104)] + [chr(1105), chr(1025)]))
    res.sheet_ru_symb_bigrams()  # additional keyword argument – ignore_case (boolean)
    # just test. There is more easier way with sheet_ru_symb_bigrams(ignore_case=True)
    res.sheet_custom_symb_bigrams(ru_symbs, ignore_case=True, name='Russian letter bigrams')
    res.sheet_yo_words()  # additional arguments – limits (int), min_quantity (int)
print('fin at:', datetime.now().strftime('%H:%M:%S'))
print('total time taked to xlsx generating:', datetime.now() - start)