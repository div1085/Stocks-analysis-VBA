
# Stock Analysis

## Overview of Project

This model aims to review the historical data of the 12 selected stocks and measure their performance over the year 2017, and 2018. 

This model is the refactor code of the original stock analysis, with the goal of improving overall processing efficiency.

### Purpose
The purpose of the model is to refactor the VBA code, with the help of looping tickers with their index (numbers), thus removing need for looking up values in each individual arrays.


## Analysis of Outcomes 

#### 2017 Data Original Code:

![2017_](https://github.com/div1085/stocks-analysis/blob/1eec9239e4551af435c89bb0125c2eb3cfe22920/Resources/original_2017.png)


#### 2017 Data with Refactored Code:


![2017_Stock_Analysis](https://github.com/div1085/stocks-analysis/blob/d60a5d22c137603653d88384f14b6bd351af7806/Resources/VBA_Challenge_2017.png)


As we can see, for year 2017, original code took 1.13 seconds to run; in refactored code, however, time significantly dropped to 0.28 seconds.


#### 2018 Data Original Code:

![2017_](https://github.com/div1085/stocks-analysis/blob/1eec9239e4551af435c89bb0125c2eb3cfe22920/Resources/original_2017.png)


#### 2018 Data with Refactored Code:

![2018_Stock_Analysis](https://github.com/div1085/stocks-analysis/blob/d60a5d22c137603653d88384f14b6bd351af7806/Resources/VBA_Challenge_2018.png)


For year 2018, original code took 1.19 seconds to run, however, refactored code only took 0.25 seconds.

From both results, it can be inferred that refactored code is almost 4 times faster than original code.


## Summary:

### Advantages and disadvantages of refactoring code:

#### Advantages of refactoring code:
	- Steps can be reduced, thus making code less complex
	- Smaller code are easier to maintain, and store
	- Can be made useful for future users of code

#### Disadvantages of refactoring code:
	- Have potential to introduce bugs in script
	- Can consume more time, thus increasing the cost and time for organization

### Advantages and disadvantages of VBA script:

Refactored VBA script did save the time in running the macro, however, that time saved is too little to be of any impact. In terms of risk-reward matrix, the reward (time saved) is too little comapred to risks (time and cost consumed, risk of corrupting the code) attached.  
