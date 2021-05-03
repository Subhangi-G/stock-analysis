# stock-analysis
# Data Analysis of Green Energy Stocks with VBA

## Overview of Project

### Purpose

The purpose of this project is to analyze the performance of different green energy stocks to help make desicions about future investments.\
The dataset contains information for 12 different stock tickers for two years, 2017 and 2018, in different worksheets.\
For each year, tickers were arranged by rows. Each row contained values of stock prices for different dates, and the volumes traded. Identical tickers were bunched together and then the different tickers were arranged alphabetically.\
The analyses provide the total of volumes traded, and the annual returns for each stock ticker for either 2017 or 2018, depending upon the choice of the user.\
The time taken for each analysis to execute is also recorded and displayed.

## Analysis

The unique stocks tickers were assigned to an array as seen below. 

```
'2) Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
```

Analysis was done by looping through the rows in the table once.\
The daily volumes for each unique ticker were added to get the total daily volume. The starting and ending values for each unique ticker was also recorded.\
Each of these information were stored in an array corresponding to the unique ticker.

```
'Get the number of rows to loop over and assign to RowCount

'Create a ticker index, and assign initial value of zero because the index of the array variables corresponding to the tickers will range from 0 through 11 (for 12 different tickers). tickerIndex = 0

'Declare three output variables arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
     
    'Create a for loop to initialize the tickerVolumes corresponding to each tickers to zero.
                 
    'Loop over all the rows in the spreadsheet from 2 through RowCount
        
       'Increase total ticker volume for current ticker, i.e. tickerVolumes(tickerIndex), by adding volume of current row
        
       'Check if the current row is the first row with the selected ticker. If True, set starting price for current ticker
             tickerStartingPrices(tickerIndex) = Value in current row
       
       'Check if the current row is the last row with the selected ticker. If True, set ending price for current ticker, and increase the ticker index by 1
             tickerEndingPrices(tickerIndex) = Value in current row
             Increase the tickerIndex by 1. This will move to the next ticker
     
      'Continue to loop to the next row
   
    'Loop through the tickers array to display ticker, their total daily volume, and return.
   
```

Time was recorded at the start and at the end of each analysis.\
The returns for each unique stock were calculated using the above information, and displayed along with their total daily volumes. The time taken for the execution is also shown.


## Results

### The stock performance for 2017
All the stocks,with the exception of TERP, gave positive returns. DQ did exceptionally well followed by ENPH and FSLR.



### The stock performance for 2018
The returns for all stocks with the exceptions of ENPH and RUN were negative. Both ENPH and RUN had similar returns.



### Comparison between the performance of the two years.

There appears to be a marked difference in the performance of the stocks between the years.\
ENPH and RUN gave positive annual returns consistently for both years, and might be considered less risky.

### Comparison between the original and refactored script

The analyses were perfomed in an earlier script by using two loops (referred to as original script). The first loop went over the different tickers. For each ticker, the nested loop went over all the rows.

```
  'Get the number of rows to loop over and store in RowCount
  
    'Loop through tickers array
                 
        'Set initial total volume to zero
                
           'Loop through rows from 2 through RowCount 
            
               'Increase totalVolume if ticker in this row is the ticker current value
                    
               'Check if the current row is the first row with the selected ticker. If True, set starting price for current tickeSet starting price for current ticker
                             
               'Check if the current row is the last row with the selected ticker. If True, set ending price for current ticker
               
           'Continue to loop to the next row
                      
           'Display current ticker value, totalVolume, and returns
           
    'Continue to loop to the next tickers array
 ```
As there are 12 different tickers, and 3013 rows. Therefore, the rows are looped over for more than 36,000 times!\ 
To shorten the time of execution, the code was refactored (see first code above) so that the looping over the of rows happened just once.\
This significantly reduced the time taken to do the analyses for both years, as seen below.


## Summary

1) Refactoring code may lead to significant reduction in time, and execute the code more efficiently. It may use less computer memory.
It may make it easier for future users of the code to understand it, or may not depending on the refactoring.

2)In these analyses refactoring the code reduced the time of execution by making the code more efficient.
In the original code the loop went over all the rows even when the the data for the ticker of interest was collected. That was unnecessary repetition of looping over the rows.
In the refactored code, the code loops over the rows only once, and collects the required data for the tickers as required and stores them in an array. 
