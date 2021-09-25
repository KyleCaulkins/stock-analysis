# Stock-Analysis

## Purpose

The purpose of this excercise was to help a friend, Steve, analyze stocks at the click of a button. Steve was interested in annual *Return* and *Total Daily Volume* of shares traded. The analysis began as a way to investigate shares of DAQO New Energy Corporation, but soon Steve wanted to expand the analysis. What started as a single share calculation expanded to a dozen shares and then Steve wanted the ability to look at a many more shares. The original code worked well for small amounts of data but may not work well for large amounts of data, or may tak a long time. The code was *refactored* in order to produce the same information but allow the VBA script to run faster. 

## Results

### Stock Analysis
![2017_vs_2018_Stocks](/Resources/2017_vs_2018_stocks.png)

Steve's parents were particularly interested in DAQO New Energy Corporation, ticker DQ, as a long term investment. DQ had a very high return for 2017; it was the highest return in the data set. 2018 was not as favorable for DQ; it had a significant loss for the year. More data year over year will need to be acquired to determine if DQ is a good investment for the long term. For the two years of data both "ENPH" and "RUN" had positive returns, both years. The analysis tool provided to Steve should help him make a better determination when more data is available. 

### VBA Script and Performance

#### Original Code
![Original_Code](/VBA_original_script.vbs)

![Original_Script_Runtime](/Resources/2018Old_Analysis_Runtime.png)

The original stock analysis is a macro designed to be run using Microsoft Excel. When run, the macro formats an output sheet with headers, uses user input year to select data to analyze, sums the total volume each ticker is traded, calculates annual return, outputs a table of total volume and annual return for each ticker, and creates a message box with the time taken to run. The code has a nested loop to sum the volume traded, the starting price, and the ending price for each ticker. The code is designed to loop over the worksheet of data to provide the outputs. The analysis takes a little over a second to run for the tweleve tickers used. 

#### Refactored Code
![Refactored_Code](/VBA_Challenge.vbs) 

![Refactored2017_Runtime](/Resources/VBA_Challenge_2017.png)

![Refactored2018_Runtime](/Resources/VBA_Challenge_2018.png)

The refactored code is also a macro designed to be run using Microsoft Excel. The refactored analysis requires the same user input for year, and then provides the same outputs for the analysis, as the original code. The primary difference in the refactored code is the use of arrays. There is an ticker index array, a ticker volumes array, a starting price array, and an ending price array. There is still a nested loop, but instead of looping through worksheet cells, the code loops through arrays. This made a significant difference in the recorded run time. The refactored code takes roughly a fifth to a quarter of the time it took to run the original code.

## Summary

In this challenge, refactoring code was the primary objective. The main driver for refactoring this code was to reduce the time taken to produce results from analysis. If the data set were very large, it would have taken the old version of the code much longer to run when compared to the refactored version. Other advantages to refactoring code could be to remove bugs, make the code easier to follow, allow for a new function, make it easier to modify, and to improve performance. All of these are great benefits, but it became obvious that there are disadvantages to refactoring code. 

In this challenge a big trade off was time. It took much longer than originally anticipated to refactor the code to run faster. Time would be a large consideration for a manager trying to prioritize work for a group of developers. Some other disadvantages could be the introduction of bugs, or other mistakes that would increase time to fix. If the analysis were only going to be used once or twice, it may not be worth the effort to refactor. The original vba script for the stock analysis ran, but after several iterations of refactoring, errors were introduced and the code would not perform the analysis. If this work had been close to a deadline, it would not have produced any progress. After more time was committed the code was eventually fixed and ran better than the original.
