# stock-analysis
# Data Analysis of Green Energy Stocks with VBA

## Overview of Project

### Purpose

The purpose of this project is to analyze the performance of different green energy stocks to help make desicions about future investments.
The dataset contains information for 12 different stock tickers for two years, 2017 and 2018, in different worksheets.\
For each year, the tickers were arranged by rows. Each row contained values of stock prices for different dates, and the volumes traded for the ticker specified in that row. Identical tickers were bunched together and then the different tickers were arranged alphabetically.\
The analyses provide the total of volumes traded, and the annual returns for each stock ticker for either 2017 or 2018, depending upon the choice of the user.\
The time taken for each analysis to execute is also recorded and displayed.

## Analysis

The unique stock tickers were assigned to an array as seen below. 

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

Analysis for the year chosen by the user, was perfomed by looping through all the rows in the table specific for that year, only once.\
The daily volumes for each unique ticker were added to get the total daily volume. The starting and ending values for each unique ticker was also recorded.\
Each of these information were stored in an array corresponding to the unique ticker.

```
'Get the number of rows to loop over and assign to RowCount

'Create a ticker index, and assign initial value of zero because the index of the array variables corresponding to the tickers will range from 0 through 11 (for 12 different tickers). tickerIndex = 0

'Declare three output variables arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
     
    'Create a for loop to initialize the tickerVolumes, corresponding to each tickers, to zero.
                 
    'Start a loop to go over all the rows in the spreadsheet from 2 through RowCount
        
       'Increase total ticker volume for current ticker (identified by the tickerIndex), stored in the array tickerVolumes(tickerIndex), by adding volume of current row
        
       'Check if the current row is the first row with the selected ticker. If it is, then set the starting price for current ticker
             tickerStartingPrices(tickerIndex) = Value in current row
       
       'Check if the current row is the last row with the selected ticker. If it is, then set the ending price for current ticker, and increase the ticker index by 1
             tickerEndingPrices(tickerIndex) = Value in current row
             Increase the tickerIndex by 1. This will move to the next ticker
     
      'Continue the loop over to the next row
   
    'Loop through the tickers array to display ticker, their total daily volume, and return.
   
```

Time was recorded at the start, and at the end of each analysis.\
The returns for each unique stock were calculated using the above information, and displayed along with their total daily volumes. The time taken for the execution is also displayed.


## Results

### The stock performance for 2017
All the stocks, with the exception of TERP, gave positive returns. DQ did exceptionally well followed by ENPH and FSLR.

![Analysis_2017](https://user-images.githubusercontent.com/71800628/116924006-ef5f6400-ac1c-11eb-8519-2132fcd09d47.png)

### The stock performance for 2018
The returns for all stocks with the exceptions of ENPH and RUN were negative. Both ENPH and RUN had similar returns.

![Analysis_2018](https://user-images.githubusercontent.com/71800628/116924054-fdad8000-ac1c-11eb-8782-51a957c6b5db.png)

### Comparison between the performance of the two years.

There appears to be a marked difference in the performance of the stocks between the years.\
ENPH and RUN gave positive annual returns consistently for both years, and might be considered less risky.

### Comparison between the original and refactored script

The analyses were perfomed in an earlier script (referred to as original script) by using two loops . The first loop went over the different tickers. For each ticker, the nested loop went over all the rows.\
Below is the pseudocode for the original script (the pseudocode for the refactored code is given above under the "Analysis" section.)

```
  'Get the number of rows to loop over and store in RowCount
  
    'Stat the loop to go through the tickers array
                 
        'Set the initial total volume to zero
                
           'Loop through rows 2 through RowCount 
            
               'Increase totalVolume if ticker in this row is the ticker current value (from the ticker array)
                    
               'Check if the current row is the first row with the selected ticker. If it is, then set the starting price for current ticker
                             
               'Check if the current row is the last row with the selected ticker. If it is, then set the ending price for current ticker
               
           'Continue the loop over to the next row
                      
           'Display current ticker value, totalVolume, and returns
           
    'Continue the loop over to the next ticker in the array
 ```
As there are 12 different tickers, and 3013 rows. Therefore, the rows are looped over for more than 36,000 times!. 
To shorten the time of execution, the code was refactored so that the looping over the of rows happened just once.\
This significantly reduced the time taken to do the analyses for both years, as seen below.

![RunTime-comparison_2017](https://user-images.githubusercontent.com/71800628/116924310-54b35500-ac1d-11eb-8250-0eebca0b4321.png)

![RunTime-comparison_2018](https://user-images.githubusercontent.com/71800628/116924345-5e3cbd00-ac1d-11eb-9d1c-203fa4294a00.png)

## Summary

1) Refactoring code may lead to significant reduction in time, and execute the code more efficiently. It may use less computer memory.
It may make it easier for future users of the code to understand it, or may not depending on the refactoring. It is therefore recommended to use comments to help deciphering codes easier\
A drawback of refactoring may be the possible introduction of bugs.

2) In these analyses refactoring the code reduced the time of execution by making the code more efficient.
In the original code, all the rows were looped over for every single ticker, even after the data for the ticker was collected. That created unnecessary repetition of looping over all the the rows.
In the refactored code, the code loops over the rows only once, and collects the required data for the tickers as required and stores them in an array. 
