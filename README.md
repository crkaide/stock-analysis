# VBA of Wall Street

## Overview of Project

### Purpose

Initial analysis was done for the client (Steve) on a dataset that he provided containing stock performance information for the DAQO New Energy Corp (DQ) and its industry competitors.  Information reported for each relevant ticker index includes the stock's performance over the timespan of the provided data as well as total volume traded (as an indicator of demand/investor buy-in).  The conclusion of the initial analysis was that the workbook initially provided was successfuly and accurately reported results, but needed to have its code refactored to allow it to run quickly over larger data sets.  Recoding goals included un-nesting For loops to force the analysis to perform just a single loop through all the data (data set should be arranged in chronological order by Column B "Date" before running the macro/analysis), as well as adding output range formatting to the analysis macro.

## Analysis and Challenges

### Analysis of Results and Code Performance

#### Stock Performance (formatted charts)

![Stock Performance 2017](https://github.com/crkaide/stock-analysis/blob/main/Results_2017.png?raw=true)

![Stock Performance 2018](https://github.com/crkaide/stock-analysis/blob/main/Results_2018.png?raw=true)

#### Code Performance (images)

##### **_REFACTORED_ Code Run Times**

![Refactored Code Run Time 2017: VBA_Challenge_2017.png](https://github.com/crkaide/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png?raw=true)

![Refactored Code Run Time 2018: VBA_Challenge_2018.png](https://github.com/crkaide/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png?raw=true)

##### **_ORIGINAL_ Code Run Times**

![Module code (original) run time, 2017, REFERENCE ONLY](https://github.com/crkaide/stock-analysis/blob/main/module%20run%20time_2017%20(reference%20only).png?raw=true)

![Module code (original) run time, 2018, REFERENCE ONLY](https://github.com/crkaide/stock-analysis/blob/main/module%20run%20time_2018%20(reference%20only).png?raw=true)

> Stock Performance:  The formatting in the "Return" columns of the Stock Performance charts shows the clearest picture of these results.  While the overwhelming majority of stocks (11 of 12) in this chart reported positive returns for 2017 (including DQ, +199.4%), the inverse is true for the following year.  In 2018, 10 of 12 stocks declined in value, with DQ showing the sharpest decline of all the stocks at -62.6%.  Despite the decline in the return, DQ traded approximately triple the volume in 2018 that it did the year before.  However, ENPH showed a similar increase in volume traded, while also reporting a +81.9% return in 2018.  Based on this data, it's recommended that investments in DQ be moved to ENPH. 

> Code Performance:  


PERCENT CHANGE/IMPROVEMENT FOR REFACTORED CODE



Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.


### Challenges and Difficulties Encountered

Only two years
Output arrays







## Summary

What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script?
