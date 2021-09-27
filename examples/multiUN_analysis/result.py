from string import ascii_lowercase

import frequency_analysis

with frequency_analysis.Result() as res:
    res.treat(limits=[1000]*4, min_quantity=[10]*5)
    res.sheet_custom_symb(ascii_lowercase)
    res.sheet_en_symb_bigrams()
