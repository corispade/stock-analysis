# Stock Analysis Using Excel VBA

Sock Analysis Challenge [dataset](https://github.com/corispade/stock-analysis/blob/main/VBA_Challenge.xlsm) and [timer references](https://github.com/corispade/stock-analysis/tree/main/Resources).

## **Overview of Project**

**Purpose:**

Steve wants to expand his dataset to include the entire stock market over the last few years. Our original code takes too long to execute for such a large amount of data. In order to reduce the execution time, we are refactoring our code to help the VBA script run faster.

## **Results**

**Process:**

We created a [tickerIndex] variable to access the correct index across the tickers, volume, starting prices and ending prices data. This reduces the number of times we are looping through the code. We continue to loop volume, starting prices and ending prices data by utilizing the [tickerIndex] variable. At the end, we looped through our arrays to output the correct data to our assigned columns. See [worksheet](https://github.com/corispade/stock-analysis/blob/main/VBA_Challenge.xlsm) for reference. 

**Analysis:**

After refactoring our code, the execution times reduced by 87.5% from 1.6 seconds average to 0.2 seconds average. See [screenshots](https://github.com/corispade/stock-analysis/tree/main/Resources) for reference. The reduced speed is much more condusive for Steve to execute code for a larger dataset.

## **Summary**

**Advantages of Refactoring Code:**

One advantage of refactoring code is cleaning up the data to make it more simplified for software to understand and in return it executes the program faster. Another advantage is that refactoring would make the code more organized and easier for future users to read and understand.

**Disadvantages of Refactoring Code:**

One distadvantage of refactoring code is that it is very time consuming to figure out how to make the code more efficient and then execute the new code properly without bugs. It could also be extremely difficult to refactor a code created by another user if it is poorly documented. 

**Advantages and Disadvantages of Original vs. Refactored VBA Script:**

The advantage of the orginial VBA script was that we were able to walk through the coding process step by step for easier understanding and application of code. This, of course, is time consuming and too complicated of a script for Excel to run efficiently.

The largest advantage of refactoring our code was evident based on the results of this project. After refactoring our code, we were able to cut the execution time by an average of 87%. Refactoring our code made the program easier to read and in return, it ran much more efficiently.

The largest disadvantage that I experienced in this project is the time it took for me to refactor the code. I came across many bugs and it took a long time to figure out the correct syntax for the code to run properly.
