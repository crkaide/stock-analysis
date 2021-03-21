# VBA of Wall Street

## Overview of Project

### Purpose

Initial analysis was done for the client (Steve) on a dataset that he provided containing stock performance information for the DAQO New Energy Corp (DQ) and its industry competitors.  Information reported for each relevant ticker index includes the stock's performance over the timespan of the provided data as well as total volume traded (an indicator of demand/investor confidence).  The conclusion of the initial analysis was that the VBA code in the first workbook successfully and accurately reported results, but needed to be refactored to allow the analysis to run quickly over larger data sets.  Recoding goals included unnesting For loops to force the analysis to perform only a single loop through all the data (**the data set should be arranged in chronological order earliest-->most recent by Column B "Date" before running the refactored macro/analysis**), as well as adding output range formatting to the analysis macro.

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

> **Stock Performance:**  The formatting in the "Return" columns of the Stock Performance charts shows the clearest picture of these results.  While the overwhelming majority of stocks (11 of 12) in this chart reported positive returns for 2017 (including DQ, +199.4%), the inverse is true for the following year.  In 2018, 10 of 12 stocks declined in value, with DQ showing the sharpest decline of all the stocks at -62.6%.  Despite the decline in the return, DQ traded approximately triple the volume in 2018 that it did the year before.  However, ENPH showed a similar increase in volume traded, while also reporting a +81.9% return in 2018.  Based on this data, it's recommended that investments in DQ be moved to ENPH. 

> **Code Performance:**
> * Refactored code 2017 ran in 0.140625 seconds.
> * Refactored code 2018 ran in 0.140625 seconds.
> * Original module code 2017 ran in 0.5625 seconds.
> * Original module code 2018 ran in 0.546875 seconds.
> 
> * **_Percent change for 2017 refactored code:  -75%_**
> * **_Percent change for 2018 refactored code:  -74%_**

**REFACTORED CODE RUNS _SIGNIFICANTLY_ FASTER**

## Summary

I encountered significant difficulties refactoring this code and am proud to have come to a final result that runs correctly after extensive research as well as trial and error.  Although I'm happy with the reformatted code, it's not ideal and not what I would hope to produce in a real-world setting.  The primary drawback of the reformatted code--and the one I tried so hard to overcome--is the hard-coding of the ticker array.  In this code, the elements of the array are explicitly defined after it's defined.  The following arguments refer back to this hard-coded array to link volumes and returns to the proper index.  The potential problem arises if this macro is run on a data set with an index list that differs from the defined array.  Ideally, the tickerIndex variable would iterate over Cells(i, 1) from 2 To RowCount, return an unduplicated list of tickers based on those values, define those tickers as a dynamic array, and use that as the tickerIndex.  I made extensive attempts to accomplish this by trying to apply the ReDim statement to various combinations of the tickerIndex and output arrays, as well as creating a For loop to iterate over the values in Column A.  Hard-coding was the only successful solution.

General advantages of refactoring code include improving its performance time (an advantage which increases with the size of the data set), improving its logic by simplifying processes and streamlining steps, and unnesting For loops that force unnecessary or inefficient reiterations of a macro's arguments and functions.  This refactoring offers all these advantages.  While the original code included a nested For loop that iterated first over the tickers and then over all the rows for each ticker, performing various conditional analyses as it does, the refactored code closes every For loop at its first "level" before beginning any additional or new iterations.

The only significant disadvantage to refactoring code is the time invested in the actual refactoring.  The larger the data set, the more time refactoring will save.  However, refactored code must produce the same results as the original, and in some cases--as in this one--any time saved by the refactored code is significantly offset by the time it takes to rewrite it correctly.

Although the percent changes reported above contradict this conclusion, in retrospect, more time was spent refactoring the code than was (or would ever be) saved by the improvements to the code.
