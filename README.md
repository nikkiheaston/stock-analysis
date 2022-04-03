# Stock Analysis in VBA
## Overview of Project
### Purpose
The purpose of this analysis is to determine which of two coding scripts in VBA is more efficient, using data for 12 specific stocks, and those stocks trade volumes and returns.  

### Results
Using stock data for the years 2017 and 2018, I analyzed 12 stocks and calculated daily volume (the total number of shares that were traded throughout a day) and the return (the difference of the ending stock price and the beginning stock price). The output of the analysis is a table in Excel displaying each stock ticker, the sum of its total daily volume, and return. Using the timer function, I compared the execution time of the first coding script against the second script. The first script that I ran completed the analysis in 0.8320313 seconds for 2017 and 0.8515625 seconds for 2018.

After refactoring the code and running the scripts again, I saw improvement in the execution time, as shown in the following screenshots:
![2017 Code Execution Time](https://github.com/nikkiheaston/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG)

![2018 Code Execution Time](https://github.com/nikkiheaston/stock-analysis/blob/main/Resources/VBA_Challenge_2018.PNG)

As we see, the code ran more than half a second faster for each year. 

### Summary
The advantage of refactoring code is that hopefully the new code will be more efficient, saving time for you as the data analyzer, and less work for your computer.  The disadvantage of refactoring code is the time that must be dedicated to writing new code, and debugging new code. If you have a script that works, and you need to spend a lot of time debugging new code that's designed to yield the same output, it's important to determine how much time you are really saving. 

Refactoring the code in this project was an advantage because it ran faster and was more efficient for the computer - a bonus if I were to extend my analsis to include more stocks or years of data. The disadvantage for me in refactoring the code was in needing to learn new ways to code which took a significantly greater amount of time than it takes to run the code in the first script.  
