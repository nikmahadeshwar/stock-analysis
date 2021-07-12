# Stock Analysis using VBA
## Overview: Project Analysis
In this project we are helping Steve for the data analysis of the green stock data set to analyze, edit and refactor the code to determine the efficiency of the code. 
We have used Excel VBA to find the stocks daily volume and Annual Return.
We don't intend to add new functionality but rather make it more effecient by reducing the memory, improving the logics and takes fewer steps.

## Purpose
In this Project, our purpose was to find an efficient way to analyse the data set of stocks using VBA. 
We were able to run the analysis on 12 different stocks and conclude that there are more effecient ways to work with the data by refactoring the code.

## Results
We improved the efficiency of the code by using for loops. We used the tickerIndex to access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays. The tickers array helped to know the ticker symbol of a stock.

## Refactored Code using for loop
1. Created a tickerIndex variable and set it equal to zero before iterating over all the rows. Used the tickerIndex to access the correct index across the four different arrays by using: the tickers array and the three output arrays.
2. Created a for loop to initialize the tickerVolumes to zero.
3. Created a loop that will loop over all the rows in the spreadsheet. Inside the loop, we created a script to increase the current tickerVolumes variable and added the ticker volume for the current stock ticker.
4. By using the if-then statement checked if the current row is the first row with the selected tickerIndex. If it is, then assigned the current starting price to the tickerStartingPrices variable.
5. By using the script that increased the tickerIndex if the next row’s ticker didnt matched the previous row’s ticker.
6. Used a for loop to loop through arrays (tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices) to output the “Ticker,” “Total Daily Volume,” and “Return” columns in your spreadsheet.
![alt text](https://github.com/nikmahadeshwar/stock-analysis/blob/main/stockanalysiscode.PNG)

## Summary
## Advantages of Refactoring Code
By refactoring code it makes our code more organised. It can be easily read, understand and maintained. It increases the quality, design and gives good results.

## Disadvantages of Refactoring Code
Refactoring code can be expensive and risky. It may also introduce bugs to the application. 

## Benifits for the current Project
The benifit we had as a result of the refactoring was decrease in macro run time. The runtime for new analysis took about 2 seconds to run. Attached below are the screenshots that indicate the run time for our new analysis.

## Stock performance of 2017
![alt text](https://github.com/nikmahadeshwar/stock-analysis/blob/main/2017%20stockanalysis.PNG)
## Stock performance of 2018
![alt text](https://github.com/nikmahadeshwar/stock-analysis/blob/main/2018%20Stock%20Anlysis.PNG)
