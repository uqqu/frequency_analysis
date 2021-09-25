import frequency_analysis

with frequency_analysis.Result() as res:
    res.treat(limits=[100]*4, min_quantity=[20]*5)
