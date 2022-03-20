# Stock Analysis VBA Challenge

## Overview
This project aims to analyze stocks for the chosen year in the excel workbook using VBA script. The original VBA code performs the desired analysis; however, refactoring the code will improve the efficiency and allow greater functionality in the future. 

## Results
### Original Code 
The original VBA code uses nested loops to examine the rows of data within each year of the stocks (2017 & 2018). This means the inner loop runs many times inside the outer loop with each iteration of the outer loop restarting the inner loop. In this coding design, the outer loop is not available to run again until the inner loop finishes all the steps. This results in a longer run time for the VBA script to complete and populate the results worksheet, “All Stocks Analysis.”  In the outer loop, the stock’s ticker is set to an index list that the outer loop runs through. Next, the inner loop will add data matching the ticker that the outer loop is currently focused on. Continuing the conditions of the inner loop, a cell value not matching the ticker will signal for the outer loop to move to the next ticker. The process is repeated until all stock tickers are read through. By refactoring, the efficiency of the code can be markedly improved. 
 
The original code ran the dataset for the years 2017 and 2018 in 1.80 seconds and 1.86 seconds respectively. 
  
### Refactored code 
To achieve a more efficient run time, the code runs through the dataset of the year selected one time by creating arrays for the information retrieved by a series of loops. The inner loops are dropped, creating a cleaner code and reducing the layers the code has to process. 
  

The refactored code ran the dataset for the years 2017 and 2018 in 0.33 seconds and 0.41 seconds respectively.
  
## Summary
Code refactoring provides the ability to restructure existing code to extend the capabilities of the code. This allows for unnecessary lines of code to be eliminated such as the nested loops from the original code for the stock analysis. A cleaner code that is more straightforward in how it is read requires less time and memory usage, resulting in efficient run times. Also, this refactored code will be easier to add other functionality if the user desires to. However, refactoring a code can be time-consuming if the original is repetitive. There is also the risk of making mistakes and additional time spent solving the problem if the refactored code does not work correctly. The complexity of the code will affect the time commitment and chances of mistakes. Despite these disadvantages, refactoring is a beneficial method for improving a code when deemed necessary. 
