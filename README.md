# Stock Analysis in VBA
## Overview of Project
### Purpose
The purpose of this analysis is to determine which of 12 specific stocks yield the highest returns and trading volume, and to determine which of two coding scripts in VBA is more efficient in completing the stock analysis.   

### Results
Using stock data for the years 2017 and 2018, I analyzed 12 stocks and calculated daily volume (the total number of shares that were traded throughout a day) and the return (the difference of the ending stock price and the beginning stock price). The output of the analysis is a table in Excel displaying each stock ticker, the sum of its total daily volume, and return. Comparing the two years, 2017 yielded positive returns for 11 of the 12 stocks, ranging from 5.5% to 199.4%. 2018 was a difficult year - only 2 stocks yielded positive returns. 

For the first script, I used a nested loop, starting with a For loop to loop through all the tickers, and then another For loop, looping through all the stock volumes, starting prices, and ending prices using If Statements. 

Example snippets of the code (see Excel file for complete code):
```
For i = 0 to 11
  
  For j = 2 to EndRow
  
  If Cells(j, 1).Value = ticker Then 
  totalVolume = totalVolume + Cells(j, 8).Value
  End If
  
  If Cells(j, 1).Value = ticker And Cells(j-1, 1).Value <> ticker Then
  startingPrice = Cells(j, 6).Value
  End If
  
  If Cells(j, 1).Value = ticker And Cells(j-1, 1).Value <> ticker Then
  endingPrice = Cells(j, 6).Value
  End If
  
  Next j
  
Next i
```

In refactoring the code, I created output arrays for the trade volumes, starting prices, and ending prices, and created a ticker index variable to access the correct index in each array. I created a For loop to initialize the trade volumes to 0. Then I created a second For loop to loop through all the rows, calculating the trade volumes, using If statements for finding the ticker starting prices and ending prices, and then advancing the ticker onto the next before looping again.  

Example snippets of the code (see Excel file for complete code):

````
'tickers array previously defined as tickers(12) 

tickerIndex = 0

tickerVolumes(12)
tickerStartingPrices(12)
tickerEndingPrices(12)


For i = 0 To 11

  tickerVolumes(i) = 0

Next i


For i = 2 to EndRow

  tickerVolumes(tickerIndex) = Cells(i, 8) + tickerVolumes(tickerIndex)

  If Cells(i-1, 1).Value <> tickers(tickerIndex) Then
  tickerStaringPrices)tickerIndex) = Cells(i, 6).Value
  End If

  If Cells(i+1, 1).Value <> tickers(tickerIndex) Then
  tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

  tickerIndex = tickerIndex + 1

  End If

Next i
````

Using the timer function, I compared the execution time of the first coding script against the second script. The first script that I ran completed the analysis in 0.8320313 seconds for 2017 and 0.8515625 seconds for 2018. 

After refactoring the code and running the scripts again, I saw improvement in the execution time, as shown in the following screenshots:


![2017 Code Execution Time](https://github.com/nikkiheaston/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG)


![2018 Code Execution Time](https://github.com/nikkiheaston/stock-analysis/blob/main/Resources/VBA_Challenge_2018.PNG)

As we see, the code ran more than half a second faster for each year. 

### Summary
The advantage of refactoring code is that hopefully the new code will be more efficient, saving time for you as the data analyzer, and less work for your computer.  The disadvantage of refactoring code is the time that must be dedicated to writing new code, and debugging new code. If you have a script that works, and you need to spend a lot of time debugging new code that's designed to yield the same output, it's important to determine how much time you are really saving. 

Refactoring the code in this project was an advantage because it ran faster and was more efficient for the computer - a bonus if I were to extend my analsis to include more stocks or years of data. The disadvantage for me in refactoring the code was in needing to learn new ways to code which took a significantly greater amount of time than it takes to run the code in the first script.  
