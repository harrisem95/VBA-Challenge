# Refactoring VBA Code for Stock Analysis

## Overview of Project

### Purpose
Steve's family is interested in investing money into stocks, but they do not have a clear understanding of how these stocks are performing. Using vba and excel, we were able to analyze the stock's total volume and return over a specific year. Calculating the total volume is important because if a stock is traded often, Steve's family believes that it's price more accurately represents the value of the stock. The return is also important because you can determine whether there was a positive return over the specified year. The purpose of this project was to introduce us to the concept of refactoring code to make it faster and more efficient. The original code we had written, although accomplished the same thing for these 12 stocks, took significantly more time. The time it takes to run the code would make a big difference if you were analyzing all of the stocks in the stock market. The purpose of this refactored code is to enable Steve to use this code to expand his research further than the 12 stocks he is currently looking into. 

## Results

### Refactoring the All Stocks Analysis VBA Code
The purpose of refactoring the code is to make it more efficient. In order to do this, I created 3 arrays to compute the total volume and the return: tickerVolumes(12), tickerStartingPrices(12), and tickerEndingPrices(12). tickerVolumes(12) held the volume, tickerStartingPrices(12) held the starting price, and tickerEndingPrices(12) held the ending price. I also created a a new variable called tickerIndex. The tickerIndex variable was used to match the ticker arrray with each array I created respectively. Using the arrays I created, I used nested for loops to loop through the data and accurately complete the analysis. Compared to my original code, this new refactored code was much more efficient. Please see images of the time it took to complete the analysis using the original code and the refactored code. You can see that my refactored code performed the analysis at a significantly faster rate.
/Users/emily/Documents/Columbia/Module 2 Challenge
#### Refactored Code Time Stamp
![Refactored Code Time Stamp 2017](/Resources/VBA_Challenge_2017.png)
![Refactored Code Time Stamp 2018](/Resources/VBA_Challenge_2018.png)

#### Original Code Time Stamp
![Original Code Time Stamp 2017](/Resources/VBA_Challenge_2017_old.png)
![Original Code Time Stamp 2018](/Resources/VBA_Challenge_2018_old.png)

## Summary

### Advantages and Disadvantages of Refactoring Code 
The advantages of refactoring code include making the code more clear and easily understood. This is imporant if someone else were to try to work on your code in the future increasing its longevitiy and uses for future projects. Refactoring the code also leads to faster programming and more efficient analysis. The disadvantage of refactoring code is you can take the original working code and introduce new bugs and errors.

### Advantages and Disadvantages of the Original and Refactored VBA Script
The advantage of the refactored VBA script is it performs the analysis at a much faster rate. This is important because if we wanted to use this script to run an analysis on a larger data set, it would be able to do so quickly compared to our original script. The disadvante of the refactored VBA Script is that although it is able to analyze larger amounts of data, we must introduce the data in the exact same format for this code to be successful. Another disadvantage to refactoring code is that it is a time consuming process. Although the end result was successful, it took me a while to work through debugging and errors that I had introduced along the way. 

