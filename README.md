# Kickstarting with Excel

## Overview of Project
The client, Steve, is a financial analyst who is requesting a VBA script that will calculate the total daily volume and the percent yearly return of different green energy stocks for a specified year.

### Purpose & Background
I will be using stock data compiled by the client and VBA to calculate the total daily value and the percent yearly return of each stock. The total daily volume is the total number of shares traded throughout the day and measures how actively a stock is traded. The yearly return percentage is the difference in price from the beginning of the year to the end of the year, this gives an idea of how much investment can grow or shrink if it is never sold. These metrics can be used to measure the success of each stock and used as evidence to support savvy investment decisions. I will not be suggesting which ticker the client should invest in as doing so is outside my area of expertise.

## Results (use images and examples of code)
I was provided a starter code to begin the stock analysis and refactored the code to fit the needs of this challenge. The runtime for each script is provided below to demonstrate the effectiveness of this refactored code.
To calculate the total daily volume of each stock, I developed a script that loops through each row of the stock data and adds up all the daily values of a specified stock to produce a total. The script then outputs this total in a separate sheet with the analysis results.
To calculate the yearly return for each year, I developed a script that loops through all the rows of stock data to find the starting and closing price of each stock for the year specified by the user. The starting and ending prices are then used to calculate the percent yearly return, which is then output to the analysis sheet.
I then used conditional formatting to illustrate the outcomes of each script and increase the overall impact of the analysis results, and created a button to run the script. The runtime of the refactored script for 2017 and 2018 is also provided below in the appropriate analysis section.

### Analysis of 2017 stock
![VBA Challenge 2017](https://github.com/mschimmy/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)
![VBA Challenge 2017 table](https://github.com/mschimmy/stock-analysis/blob/main/Resources/VBA_Challenge_2017_table.png)
From the table above, we can see that the TERP ticker was the only stock that had a negative yearly return, while all others were positive. The four tickers with the highest percent yearly return were DQ, SEDG, ENPH, and FSLR in that order. Investments not sold in each of these stocks more than doubled in 2017. The four tickers with the highest total daily value were SPWR, FSLR, CSIQ, and RUN, meaning that these stocks were the most actively traded.
We can see that the refactored code runs the script for 2017 in 0.1601562 seconds, while the original code ran in 0.8515625 seconds.

### Analysis of 2018 Stock
![VBA Challenge 2018](https://github.com/mschimmy/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)
![VBA Challenge 2018 table](https://github.com/mschimmy/stock-analysis/blob/main/Resources/VBA_Challenge_2018_table.png)
From the table above, we can see that there were only two tickers with a positive percent yearly return: RUN and ENPH. Both RUN and ENPH had a percent yearly increase of more than 80%, making them quite successful. All other tickers had a negative percent yearly return, meaning any investments not sold in these stocks lost value, with the JKS ticker losing the most, more than half its value. The four tickers with the highest total daily value were ENPH, SPWR, RUN, and FSLR.
The refactored code runs the script for 2018 in 0.1445312 seconds, while the original code ran in 0.8046875.

###Comparing Stock Performance
From our tables and analysis above, we can see that from 2017 to 2018, the only two tickers that had a positive yearly return were RUN and ENPH. Investments not sold in this stock grew in value over two years. The ticker TERP was the only ticker that had a negative percent yearly return over two years, meaning investments not sold in this stock decreased in value in both years.
The runtime of the refactored code is approximately five times faster than the original code. 

## Summary
- What are the advantages or disadvantages of refactoring code?
  -- An advantage of refactoring code is that it decreases the total amount of time and effort spent on creating the script, saving the analyst valuable resources. It allows us to not have to reinvent the wheel, and instead build upon code that has already been written. A disadvantage of refactoring code is it creates a reliance on comments from the previous analyst to understand what the code is doing. Without these comments, the analyst could spend more time trying to understand the code than what it might have taken to start from scratch. Either way, the ease or difficulty of refactoring code depends on the previous coder.
- How do these pros and cons apply to refactoring the original VBA script?
  -- Refactoring the original VBA code was relatively simple in that the provided code was free of syntax errors and structured using comments that were easy to follow. If the code had been filled with errors and without comments, it would have been much more difficult to refactor.
