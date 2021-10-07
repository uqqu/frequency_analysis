# Frequency analysis

Python package for symbol/word and their bigrams frequency analysis with excel output.
<br>What values can be counted: quantity, quantity in the first position, quantity in the last position, average position.
<br>For which data values can be counted: symbols, symbol bigrams, words, word bigrams.
<br>Additional possible data: ye-yo words table for Russian language (in excel output it can be cross-referenced with words quantity).

## Usage

1. <code>pip install frequency-analysis</code>;
2. Download data set of your choice;
3. Call <code>Analysis</code> class with context manager (take a look at the [optional arguments](#analysis-class-arguments));
4. Parse your data set to word list (one sentence in list for properly word position counting) and send it to one of three [methods of <code>Analysis</code>](#analysis-methods);
5. Call <code>Result</code> class with context manager (with optional <code>name</code> argument);
6. Call one or several of [<code>Result</code> methods](#result-methods) to create excel sheet(s) with appropriate data.

## Methods and arguments

### Analysis class arguments

> All arguments are optional

* *name* – the name of the folder in which the analysis will be saved
<br>default <code>frequency_analysis</code>
* *mode* – analysis operation mode – [n]ew, [a]ppend (to existing), [c]ontinue (to send the previous unfinished set to existing analysis)
<br>default <code>n</code>
* *word_pattern* – regex pattern for matching inwords symbols
<br>default <code>'[a-zA-Zа-яА-ЯёЁ]+(?:(?:-?[a-zA-Zа-яА-ЯёЁ]+)+|\'?[a-zA-Zа-яА-ЯёЁ]+)|[a-zA-Zа-яА-ЯёЁ]'</code>
* *allowed_symbols* – string of symbols or list with symbol unicode decimal values, which will be counted to analysis 
<br>default <code>[\*range(32, 127), 1025, \*range(1040, 1104), 1105]</code> (base punctuation, base Latin, Russian Cyrillic)
* *yo* – boolean for additional Russian word processing – compare words with word list to detect number of ye/yo misspelling.
<br>To use this mode you should place two word files near the running script (<code>yo.txt</code> for words with mandatory <code>yo</code> and <code>ye-yo.txt</code> for possibly <code>yo</code> writing). You can use your own or take it [here](https://github.com/uqqu/yo_dict).
<br>default <code>False</code>

### Analysis class methods

#### count_symbols(word_list: list, [pos: bool, bigrams: bool])
Method for counting symbol and symbol_bigram frequency.
<br>Counted values: quantity, quantity in the first position, quantity in the last position, average position in word.
<br>Average position counted only with argument <code>pos</code> as <code>True</code> (default <code>False</code>). 
<br>Position for symbols, which matched with <code>word_pattern</code> counted as for "clear" word, for other – as for "raw".
<br>Example: in single word "–Yes!" with default <code>word_pattern</code> positions will be counted as (– 1), (Y 1), (e 2), (s 3), (! 5).
<br>Bigrams counting can be disabled with argument <code>bigram</code> as <code>False</code> (default <code>True</code>).

#### count_words(word_list: list, [pos: bool, bigrams: bool])
Method for counting word and word_bigrams frequency. 
<br>Counted values: quantity, quantity in the first position, quantity in the last position, average position in sentence.
<br>Average position counted only with argument <code>pos</code> as <code>True</code> (default <code>False</code>). 
<br>Bigrams counting can be disabled with argument <code>bigram</code> as <code>False</code> (default <code>True</code>).

#### count_all(word_list: list, [pos: bool, symbol_bigrams: bool, word_bigrams: bool])
Combined call of previous two methods.

### Result class arguments

The only argument is optional
* *name* – the name of the folder in which the analysis was saved
<br>default <code>frequency_analysis</code>

### Result class methods

> First 6 methods can be called all it once with <a href="#treatlimits-tuplefour-int-min_quantities-tuplefive-int-chart_limits-tuplefour-int">treat() method</a>

Many methods accept arguments <code>limit</code>, <code>chart_limit</code>, <code>min_quantity</code> and <code>ignore_case</code>.
* *limit* (default <code>0</code>) it is a max number of elements, which will be added to the sheet. Zero – unlimited;
* *chart_limit* (default <code>20</code>) – a number of elements, which will be counted with graphical chart;
* *min_quantity* (default <code>1</code>) – a minimal appropriate value at with element will be added to the sheet;
* *ignore_case* (default <code>False</code>) – with this argument as <code>True</code> lower- and upper- case symbols will be united into a single element. With <code>False</code> – will be counted separately.

#### sheet_stats()
Main result info – number of entries, total count and average position (if exists) for each data type.

#### sheet_top_symbols([limit, chart_limit, min_quantity])
Top list of all analyzed symbols sorted by quantity. The next to it is also located the same one list, but with ignore-case. There is no need to create separate sheet, just use column of your choice.

#### sheet_top_symbol_bigrams([limit, chart_limit, min_quantity])
Top list of symbol bigrams sorted by quantity with additional <code>case insensitive</code> top-list.

#### sheet_all_symbol_bigrams([min_quantity, ignore_case])
2D sheet with all bigrams quantity. <code>min_quantity</code> argument works here for sum of row/column values instead of each separated bigram.

#### sheet_top_words([limit, chart_limit, min_quantity])
Top list of analyzed words sorted by quantity. Word counting is always case insensitive, on the <code>Analyze</code> stage.

#### sheet_top_word_bigrams([limit, chart_limit, min_quantity])
Top list of analyzed word bigrams sorted by quantity.

#### treat([limits: tuple(four int), chart_limits: tuple(four int), min_quantities: tuple(five int)])
Single call of all Result methods above. Calling methods in order of tuple values: <br><code>sheet_top_symbols()</code>, <br><code>sheet_top_symbol_bigrams()</code>, <br><code>sheet_top_words()</code>, <br><code>sheet_top_word_bigrams()</code>, <br><code>sheet_all_symbol_bigrams()</code>
<br>Please note – the last one (value for *sheet_all_symbol_bigrams()*) there is only in the *min_quantities* argument.
<br>Default values as elsewhere: *limits* – (0,)\*4; *chart_limits* – (20,)\*4; *min_quantities* – (1,)\*5.

#### sheet_custom_symbols(symbols: str, [chart_limit, name='Custom symbols'])
Create symbols top-list as <code>sheet_top_symbols()</code>, but only with symbols of your choice.

#### sheet_custom_symbol_bigrams(symbols: str, [ignore_case, name='Custom symbol bigrams'])
Create symbol bigrmas 2D sheet as <code>sheet_all_symbol_bigrams()</code>, but only with symbols of your choice.
<br>Order of symbols on the sheet will be the same as in the input argument.

#### sheet_en_symbol_bigrams([ignore_case])
Create symbol bigrams 2D sheet as <code>sheet_all_symbol_bigrams()</code>, but only with base Latin symbols.

#### sheet_ru_symbol_bigrams([ignore_case])
Create symbol bigrams 2D sheet as <code>sheet_all_symbol_bigrams()</code>, but only with Russian Cyrillic symbols.

#### sheet_yo_words([limit, min_quantity])
Create cross-referenced sheet for all counted ye-yo words with their quantity and total misspells counter. Works only with analysis created with <code>yo</code> argument as <code>True</code>.

## Performed analyses

* English analysis with [EuroMatrixPlus/MultiUN](http://www.euromatrixplus.net/multi-un/) English data set (2.9Gb .xml, 2.4\*10<sup>9</sup> symbols, 379\*10<sup>6</sup> words)
	* https://github.com/uqqu/frequency_analysis/tree/master/examples/multiUN_analysis
* Russian analysis with [OpenCorpora](http://opencorpora.org/) data set (520Mb .xml, 11.7\*10<sup>6</sup> symbols, 1.6\*10<sup>6</sup> words)
	* https://github.com/uqqu/frequency_analysis/tree/master/examples/annot_opcorpora_analysis
