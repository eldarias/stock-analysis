# Module 2 - Stock Analysis with VBA in Excel

## Overview of Project
The purpose of this project was to analyze the provided stocks by using VBA in Excel to determine the overall performance of each stock for the years **2017** and **2018**.  The analysis of the dataset was performed using VBA in Excel by putting into practice most of the already learned/reviewed VBA code functions.  The challenge was to repurpose the provided VBA code and enhance the runtime performance of the code.

## Results
The results from the project were excellent since a significant decrease on the total runtime of the code was observed, yet achieving the same results, which was to analyze the stocks' performance for each year.  As for the stocks' performance between 2017 and 2018, the analyzed stocks performed very well in 2017 where most of them had very good returns, while the 2018 stocks had bad performance with negative returns, with the exception of two stocks.  The images below provide a visualization of the runtime as well as the performance for both years.

### All Stocks for 2017

#### Original VBA Code Runtime
- <image src="./Resources/VBA_Challenge_2017-OriginalCodeRuntime.png">
  - The above image displays the total VBA Code runtime when analyzing all **2017** Stocks 'BEFORE' refactoring the VBA code.

#### Refactored VBA Code Runtime
- <image src="./Resources/VBA_Challenge_2017.png">
  - The above image displays the total VBA Code runtime when analyzing all **2017** Stocks 'AFTER' refactoring the VBA code.

### All Stocks for 2018

#### Original VBA Code Runtime
- <image src="./Resources/VBA_Challenge_2018-OriginalCodeRuntime.png">
  - The above image displays the total VBA Code runtime when analyzing all **2018** Stocks 'BEFORE' refactoring the VBA code.

#### Refactored VBA Code Runtime
- <image src="./Resources/VBA_Challenge_2018.png">
  - The above image displays the total VBA Code runtime when analyzing all **2018** Stocks  'AFTER' refactoring the VBA code.


### Stocks' performance for 2017 and 2018
#### Performance of Stocks for 2017
- <image src="./Resources/VBA_Challenge_2017-StocksPerformance.png">
  - The above image displays the performance of all analyzed stocks for the Year 2017
  
#### Performance of Stocks for 2018
- <image src="./Resources/VBA_Challenge_2018-StocksPerformance.png">
  - The above image displays the performance of all analyzed stocks for the Year 2018


### Refactored Code:
The refactored code can be obtained from here: [VBA_Challenge.vbs](https://github.com/eldarias/stock-analysis/blob/main/VBA_Challenge.vbs).

## Summary
In summary, although the end results were the same, which was to obtain the performance of all stocks for each year, the VBA code was refactored/changed to enhace its performance as well as used more advanced VBA coding methods/features.  There are advantages and disadvantages that can be said whenof refactoring the code.  One of the many advantages of refactoring the code were clearly observed by a significant decrease on the code's runtime.  This enhancement to the performance will also enable us to potentially reuse the code to analyze datasets with much higher amounts of data.  The runtime would be more noticeable and appreciated if analyzing larger amounts of data or when using systems with low system resources.  The only general disadvantage that I can see of refactoring the code is if the code will not be reused to analyze higher amounts of data due to the time spent refactoring the code since the end results were the same, but overall, I see more advantages then disadvantages.

When comparing the original and the refactored code, we can clearly see advantages on the refactored code.  For example, advantages of the refactored code is the use of output arrays and additional loops, which allows the collection of the data in less loops hence the runtime decrease/code efficiency.  The advantages mentioned for the refactored code are the disadvantage to the original code, which it takes more loops when collecting and analyzing the required data, hence the longer runtime and may be unefficient when analyzing large datasets and/or when using systems with low resources.