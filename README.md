# Stock Analysis

## Overview of Project

The project was carried out to determine profitability rate of investing in stocks of different companies. The analysis was conducted over a dozen companies for 2017 and 2018. Although the initial code, provided to the investors, works well for a dozen stocks, it might not work so well for thousand stocks. Click here to view the dataset: [VBA_Challenge.xlsm](https://github.com/sharof17/stock-analysis/blob/e6cd3762c5403bdcb0d85f98397b6a9027a3be7d/VBA_Challenge.xlsm)

### Purpose
The main purpose of the project is to refactor the code and increase its efficiency to make it work faster.

## Results
Analysis showed that refactoring the code made the it work faster. Now, investors can bigger dataset more efficiently. The following pictures demonstrate the processing time (after refactoring). 

All Stocks (2017)

![VBA Challenge 2017](Resources/VBA_Challenge_2017.png)

All Stocks (2018)

![VBA Challenge 2018](Resources/VBA_Challenge_2018.png)

In the refactored code, there wasn't used nested loop, which is for sure shortened the processing time. And, instead of generating and illustrating output one by one, it was decided, first to generate arrays with the needed outputs and then to illustrate them in the worksheet (using separate "for" loop). Click here to view the code: [VBA_Challenge.vba](https://github.com/sharof17/stock-analysis/blob/e6cd3762c5403bdcb0d85f98397b6a9027a3be7d/VBA_Challenge.vbs)

## Summary

### Pros and Cons of Refactoring Code

Refactoring leads to better quality code making it cleaner and more organized. It makes the code easier to read/understand, to find bugs, to process faster. And, one of the main disadvantage of refactoring is that it may require to retest lots of functionality. It is especially risky when application is big.

### Pros and Cons of the Original and Refactored VBA Script

The biggest advantage of refactoring is the decrease of processing time. However, original code was working well illustrating all the outputs one by one.

The disadvantage of refactoring is that it was needed to rewrite and retest "for" loops.
