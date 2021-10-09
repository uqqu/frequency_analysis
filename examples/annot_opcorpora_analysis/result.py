from datetime import datetime
import frequency_analysis

start = datetime.now()

with frequency_analysis.Result() as res:
    # treat() calls all main sheet functions.
    #   limits – max number of items for sheets (symbols, symbol bigrams, words, word bigrams)
    #   chart_limits – number of first n items for pie charts counting
    #       (symbols, symbol bigrams, words, word bigrams) sheets
    #   min_quantities – min quantity of each item for sheets
    #       (symbols, symbol bigrams, words, word bigrams, symbol bigrams 2D)
    res.treat(limits=(1000,) * 4, chart_limits=(20,) * 4, min_quantities=(10,) * 5)
    res.sheet_ru_top_symbols()  # optional argument – chart_limit (int)
    res.sheet_ru_symbol_bigrams()  # optional keyword argument – ignore_case (boolean)
    # just test. There is more easier way with sheet_ru_symb_bigrams(ignore_case=True)
    ru_symbs = 'абвгдеёжзийклмнопрстуфхцчшщьыъэюя'
    res.sheet_custom_symbol_bigrams(ru_symbs, ignore_case=True, name='Russian letter bigrams')
    res.sheet_yo_words()  # optional arguments – limit (int), min_quantity (int)
print('fin at:', datetime.now().strftime('%H:%M:%S'))
print('total time taked to xlsx generating:', datetime.now() - start)
