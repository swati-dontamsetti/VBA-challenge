# The VBA of Wall Street

## Background
VBA stands for Visual Basics for Applications. It is a programming language for the Microsoft Office of applications, like Excel. In this assignment I will use VBA scripting to analyze real stock market data. The joy of VBA is to take the tediousness out of repetitive task and run over and over again with a click of the button.

### Stock Market Analyst
![stock market](Images/stockmarket.jpg)

## Main Objective
* Create a script that will loop through all the stocks for one year for each run and take the following information:

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

* I also have conditional formatting that will highlight positive change in green and negative change in red.

* The result will look similar to this:
![moderate_solution](Images/moderate_solution.png)

### CHALLENGES

1. The solution will also be able to return the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume". The solution will look similar to this:

![hard_solution](Images/hard_solution.png)

2. I made the appropriate adjustments to the VBA script that will allow it to run on every worksheet, i.e., every year, just by running the VBA script once.

## Submission

* I have uploaded the following to Github:

  * A screen shot for each year of the results on the Multi Year Stock Data.

  * VBA Script as a separate file.
