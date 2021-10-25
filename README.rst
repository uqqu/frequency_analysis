Frequency analysis
------------------

Python package for symbol/word and their bigrams frequency analysis with excel output.

What values can be counted: quantity, quantity in the first position, quantity in the last position, average position.

For which data values can be counted: symbols, symbol bigrams, words, word bigrams.

Additional possible data: ye-yo words table for Russian language (in excel output it can be cross-referenced with words quantity).

Usage
-----

1. ``pip install frequency-analysis``;
2. Download data set of your choice;
3. Call ``Analysis`` class with context manager (take a look at the optional arguments);
4. Parse your data set to word list (one sentence in list for properly word position counting) and send it to one of three methods of ``Analysis``;
5. Call ``Result`` class with context manager (with optional name argument);
6. Call one or several of ``Result`` methods to create excel sheet(s) with appropriate data.

Methods and arguments
---------------------

``Analysis`` class arguments
~~~~~~~~~~~~~~~~~~~~~~~~~~~~

All arguments are optional

* *name* – the name of the folder in which the analysis will be saved
     default ``frequency_analysis``
* *mode* – analysis operation mode (``[n]ew``, ``[a]ppend``, ``[c]ontinue``)
     default ``n``
* *word\_pattern* – regex pattern for matching inwords symbols
    default ``[a-zA-Zà-ÿÀ-ß¸¨]+(?:(?:-?[a-zA-Zà-ÿÀ-ß¸¨]+)+\|'?[a-zA-Zà-ÿÀ-ß¸¨]+)\|[a-zA-Zà-ÿÀ-ß¸¨]``
* *allowed\_symbols* – string of symbols or list with symbol unicode decimal values, which will be counted to analysis
    default ``[*range(32, 127), 1025, *range(1040, 1104), 1105]`` (base punctuation, base Latin, Russian Cyrillic)
* *yo* – int for additional Russian word processing – compare words with word list to detect number of ye/yo misspelling. 0 – disabled; 1 – enabled; 2 with 'a' mode – update yo list with new data.
     default ``0``

     To use the last one you should place two word files near the running script (``yo.txt`` for words with mandatory yo and ``ye-yo.txt`` for possibly yo writing). You can use your own or take it `here <https://github.com/uqqu/yo_dict>`__.

``Analysis`` class methods
~~~~~~~~~~~~~~~~~~~~~~~~~~

``count_symbols(word_list: list, [pos: bool, bigrams: bool])``
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Method for counting symbol and symbol\_bigram frequency. Counted values:
quantity, quantity in the first position, quantity in the last position, average position in word. 

Average position counted only with argument ``pos`` as ``True`` (default ``False``). Position for symbols, which matched with ``word_pattern`` counted as for "clear" word, for other – as for "raw".

Example: in single word ``–Yes!`` with default ``word_pattern`` positions will be counted as ``(– 1), (Y 1), (e 2), (s 3), (! 5)``.

Bigrams counting can be disabled with argument ``bigram`` as ``False`` (default ``True``).

``count_words(word_list: list, [pos: bool, bigrams: bool])``
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Method for counting word and word\_bigrams frequency. Counted values:
quantity, quantity in the first position, quantity in the last position, average position in sentence. 

Average position counted only with argument ``pos`` as ``True`` (default ``False``).

Bigrams counting can be disabled with argument ``bigram`` as ``False`` (default ``True``).

``count_all(word_list: list, [pos: bool, symbol_bigrams: bool, word_bigrams: bool])``
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Combined call of previous two methods.

``Result`` class arguments
~~~~~~~~~~~~~~~~~~~~~~~~~~

The only argument is optional

* *name* – the name of the folder in which the analysis was saved
    default ``frequency_analysis``

``Result`` class methods
~~~~~~~~~~~~~~~~~~~~~~~~

    First 6 methods can be called all it once with ``treat()`` method

Many methods accept arguments ``limit``, ``chart_limit``, ``min_quantity`` and ``ignore_case``.

* *limit* (default ``0``) it is a max number of elements, which will be added to the sheet. ``0`` – unlimited;
* *chart_limit* (default ``20``) – a number of elements, which will be counted with graphical chart;
* *min_quantity* (default ``1``) – a minimal appropriate value at with element will be added to the sheet;
* *ignore_case* (default ``False``) – with this argument as ``True`` lower- and upper- case symbols will be united into a single element. With ``False`` – will be counted separately. ``Keyword-only``

``sheet_stats()``
^^^^^^^^^^^^^^^^^

Main result info – number of entries, total count and average position (if exists) for each data type.

``sheet_top_symbols([limit, chart_limit, min_quantity])``
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Top list of all analyzed symbols sorted by quantity. The next to it is also located the same one list, but with ignore-case. There is no need to create separate sheet, just use column of your choice.

``sheet_top_symbol_bigrams([limit, chart_limit, min_quantity])``
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Top list of symbol bigrams sorted by quantity with additional case insensitive top-list.

``sheet_all_symbol_bigrams([min_quantity, ignore_case])``
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

2D sheet with all bigrams quantity. ``min_quantity`` argument works here for sum of row/column values instead of each separated bigram.

``sheet_top_words([limit, chart_limit, min_quantity])``
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Top list of analyzed words sorted by quantity. Word counting is always case insensitive, on the ``Analyze`` stage.

``sheet_top_word_bigrams([limit, chart_limit, min_quantity])``
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Top list of analyzed word bigrams sorted by quantity.

``treat([limits: tuple(four int), chart_limits: tuple(four int), min_quantities: tuple(five int)])``
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Single call of all ``Result`` methods above. Calling methods in order of tuple values:

1. ``sheet_top_symbols()``
2. ``sheet_top_symbol_bigrams()``
3. ``sheet_top_words()``
4. ``sheet_top_word_bigrams()``
5. ``sheet_all_symbol_bigrams()``

Please note – the last one (value for ``sheet_all_symbol_bigrams()``) there is only in the ``min_quantities`` argument. 

Default values as elsewhere:

* *limits* – ``(0,)*4``
* *chart_limits* – ``(20,)*4``
* *min_quantities* – ``(1,)*5``

``sheet_custom_top_symbols(symbols: str, [chart_limit, name='Custom symbols'])``
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Create symbols top-list as ``sheet_top_symbols()``, but only with symbols of your choice. ``name`` – ``keyword-only``

``sheet_en_top_symbols(symbols: str, [chart_limit])``
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Create symbols top-list as ``sheet_top_symbols()``, but only with base Latin symbols.

``sheet_ru_top_symbols(symbols: str, [chart_limit])``
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Create symbols top-list as ``sheet_top_symbols()``, but only with Russian Cyrillic symbols.

``sheet_custom_symbol_bigrams(symbols: str, [ignore_case, name='Custom symbol bigrams'])``
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Create symbol bigrmas 2D sheet as ``sheet_all_symbol_bigrams()``, but only with symbols of your choice. Order of symbols on the sheet will be the same as in the input argument. ``name`` – ``keyword-only``

``sheet_en_symbol_bigrams([ignore_case])``
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Create symbol bigrams 2D sheet as ``sheet_all_symbol_bigrams()``, but only with base Latin symbols.

``sheet_ru_symbol_bigrams([ignore_case])``
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Create symbol bigrams 2D sheet as ``sheet_all_symbol_bigrams()``, but only with Russian Cyrillic symbols.

``sheet_yo_words([limit, min_quantity])``
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Create cross-referenced sheet for all counted ye-yo words with their quantity and total misspells counter. Works only with analysis created with ``yo`` argument as ``1`` or ``2``.

Performed analyses
------------------

* English analysis with `EuroMatrixPlus/MultiUN <http://www.euromatrixplus.net/multi-un/>`__ English data set (3.1Gb .xml, 2.4\*10\ :sup:`9` symbols, 379\*10\ :sup:`6` words)

   * https://github.com/uqqu/frequency\_analysis/tree/master/examples/en/multiUN

* Russian analysis with `EuroMatrixPlus/MultiUN <http://www.euromatrixplus.net/multi-un/>`__ Russian data set (4.3Gb .xml, 2.2\*10\ :sup:`9` symbols, 270\*10\ :sup:`6` words)

   * https://github.com/uqqu/frequency\_analysis/tree/master/examples/ru/multiUN

* Russian analysis with `OpenCorpora <http://opencorpora.org/>`__ data set (528Mb .xml, 11.7\*10\ :sup:`6` symbols, 1.6\*10\ :sup:`6` words)

   * https://github.com/uqqu/frequency\_analysis/tree/master/examples/ru/annot\_opcorpora
