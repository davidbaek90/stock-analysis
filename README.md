# stock-analysis

## Overview of Project

### Purpose

#### The purpose of this project is make use of a refactor method in provided VBA in order to process and analyze the given stock data for year 2017 and 2018. The refactoring method should provide an edge over the original code in terms of macro run speed.

## Results

#### The result of the new refactoring code provides the similar result to the original code in terms of data analysis. What is noticeability different is the time it takes for the code to execute. The original code resulted in exuction time of 0.3515 and 0.3438 seconds for year 2017 and 2018 respectively. The next two pictures show the result of running the refactored macros for the same years. Even with the same datasets, the refactored code excuted the macro in 0.0548 and 0.0625 seconds for the year 2017 and 2018. In case of stock performance of the two years, 2017 stock result outperformed year 2018 greatly, with only 1 stock falling below zero as opposed to only two stocks making the positive return in 2018. The refactored script for the VBA Challenge can be downloaded below:

[You can download my VBA Challenge here](https://github.com/davidbaek90/stock-analysis/raw/main/VBA_Challenge.xlsm)

![VBA Challenge 2017 Refactorered](https://github.com/davidbaek90/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)
![VBA Challenge 2018 Refactorered](https://github.com/davidbaek90/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

The part of the code that input as an answer to this challenge is shown in drowndown below (code in summary):

<details>
  
***<summary>  
Code in summary***
</summary>
  
    ''''
    1a) Create a ticker Index
    Dim tickerIndex As Single
    
    tickerIndex = 0
    
    '1b) Create three output arrays
    Dim tickerVolumes(11) As Long
    Dim tickerStartingPrices(11) As Single
    Dim tickerEndingPrices(11) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
    ''2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
        
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row's ticker doesn't match, increase the tickerIndex.
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
       
        '3d) Increase the tickerIndex.
        tickerIndex = tickerIndex + 1
                      
        End If
    Next i
   
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
            
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    ''''
</details>

  
## Summary

### What are the advantages or disadvantages of refactoring code?

#### The result of refactoring code is a clean, organized layout of your final code. Cleaner code helps other coders easier to understand how the code works. This is an advantage to most, if not all of the programmers. Since the code can be picked up quickly, it enables fast debugging. The pro would truly shine in a scrum environment where you are always in a time crunch to get tasks done in given time. One disadvantage of refactoring might be that since it requires more time to develop and get seasoned with programming skills, it takes time to master. For beginners like us to refactor a complex code it might take longer than to simply debug and get the problem solved. Another issue of refactoring could be that if the code is not complex enough, it might not be worth the time investment to refactor the code.

### How do these pros and cons apply to refactoring the original VBA script?

#### The most prominent advantage that was observed when the original script was refactored was the macro run time. Compared to the original script, the new code ran almost six times faster. Considering it is rather a simple script, refactoring larger script would save you more time. This would be helpful in applications where time in seconds, or even miliseconds are crucial factor in evaluating performance.
