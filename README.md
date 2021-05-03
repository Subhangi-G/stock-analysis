# stock-analysis
# Data Analysis of Green Energy Stocks with VBA

## Overview of Project

### Purpose

The Purpose of this project is to analyze the performance of different green energy stocks to help make desicions about future investments.\
The dataset contains information for 12 different stock tickers for two years, 2017 and 2018, in different worksheets.\
For each year, tickers were arranged by rows. Each row contained values of stock prices for different dates, and the volume traded. Identical tickers were bunched together and then the different tickers were arranged alphabetically.\
The analysis provides the total of volumes traded, and the annual returns for each stock ticker for either 2017 or 2018, depending upon the choice of the user.\

## Analysis

The unique stocks tickers were assigned to an array. Analysis was done by looping over the rows in the table once.\
The daily volumes for each unique ticker were added to get the total daily volume. The starting and ending values for each unique ticker was also recorded.\
Each of these information were stored in an array corresponding to the unique ticker.\
The total daily volumes and the returns were then displayed for each unique stocks.

## Results

### The stock performance for 2017



All the stocks,with the exception of TERP, gave positive returns. DQ did exceptionally well followed by ENPH and FSLR.

### The stock performance for 2018



The returns for all stocks with the exceptions of ENPH and RUN were negative. Both ENPH and RUN had similar returns.\

### Comparison between the performance of the two years.

There appears to be a marked difference in the performance of the stocks between the years.\
ENPH and RUN gave positive annual returns consistently for both years, and might be considered less risky.

### Comparison between the original and refactored script

The analyses were perfomed in an earlier script by using two loops. The first loop went over the different tickers. The nested loop went over the rows.\
As there are 12 different tickers, and 3013 rows, the rows are looped over for more than 36,000 times!\ 
To shorten the time of execution, the code was refactored so that the looping over the of rows happened just once.\
This significantly reduced the time taken to do the analyses for both years, as seen below.


## Summary

1) Refactoring code may lead to significant reduction in time, and execute the code more efficiently. It may use less computer memory.
It may make it easier for future users of the code to understand it, or may not depending on the refactoring.

2)In these analyses refactoring the code reduced the time of execution by making the code more efficient.
In the original code the loop went over all the rows even when the the data for the ticker of interest was collected. That was unnecessary repetition of looping over the rows.
In the refactored code, the code loops over the rows only once, and collects the required data for the tickers as required and stores them in an array. 
