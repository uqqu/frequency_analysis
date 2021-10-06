from datetime import datetime
from string import ascii_letters

import frequency_analysis

start = datetime.now()
with frequency_analysis.Result() as res:
    # treat() calls all main sheet functions.
    #   limits – max number of items for sheets (symbols, symbol bigrams, words, word bigrams)
    #   min_quantity – min quantity of each item for sheets
    #       (symbols, symbol bigrams, words, word bigrams, symbol bigrams 2D)
    #   chart_limit – number of first n items for charts on
    #       (symbols, symbol bigrams, words, word bigrams) sheets
    res.treat(limits=(1000,) * 4, min_quantities=(10,) * 5, chart_limits=(20,) * 4)
    # sheet_custom_symb(symbols: string/list/tuple/set, chart_limit: int, name: str (sheet_name)
    res.sheet_custom_top_symbols(ascii_letters)
    res.sheet_en_symbol_bigrams()  # additional keyword argument – ignore_case (boolean)
    # just test. There is more easier way with sheet_en_symb_bigrams(ignore_case=True)
    res.sheet_custom_symbol_bigrams(ascii_letters, ignore_case=True, name='English letter bigrams')
print('fin at:', datetime.now().strftime('%H:%M:%S'))
print('total time taked to xlsx generating:', datetime.now() - start)