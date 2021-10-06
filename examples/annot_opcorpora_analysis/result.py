from datetime import datetime
import frequency_analysis

start = datetime.now()
ru_symbs = 'абвгдеёжзийклмнопрстуфхцчшщьыъэюя'

with frequency_analysis.Result() as res:
    # treat() calls all main sheet functions.
    #   limits – max number of items for sheets (symbols, symbol bigrams, words, word bigrams)
    #   min_quantities – min quantity of each item for sheets
    #       (symbols, symbol bigrams, words, word bigrams, symbol bigrams 2D)
    #   chart_limits – number of first n items for pie charts counting
    #       (symbols, symbol bigrams, words, word bigrams) sheets
    res.treat(limits=(1000,) * 4, min_quantities=(10,) * 5, chart_limits=(20,) * 4)
    # sheet_custom_symb(symbols: string/list/tuple/set, chart_limit: int, name: str (sheet_name)
    res.sheet_custom_top_symbols(
        ''.join([chr(x) for x in range(1040, 1104)] + [chr(1105), chr(1025)])
    )
    res.sheet_ru_symbol_bigrams()  # additional keyword argument – ignore_case (boolean)
    # just test. There is more easier way with sheet_ru_symb_bigrams(ignore_case=True)
    res.sheet_custom_symbol_bigrams(ru_symbs, ignore_case=True, name='Russian letter bigrams')
    res.sheet_yo_words()  # additional arguments – limit (int), min_quantity (int)
print('fin at:', datetime.now().strftime('%H:%M:%S'))
print('total time taked to xlsx generating:', datetime.now() - start)
