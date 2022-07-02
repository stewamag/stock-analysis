# Stock_Analysis

## Overview of Project:

Data on the performance of the stocks of green energy companies was analyized using VBA and excel. The stocks of twelve companies were analyzed in this project. Their tickers were AY, CSIQ, DQ, ENPH, FSLR, HASI, JKS, RUN, SEDG, SPWR, TERP, and VSLR.

### Purpose

The data was compiled from the years 2017 and 2018 and used to determine whether each company was worth investing in based on their yearly performance.

## Analysis and Challenges

### Analysis of Stocks in 2017

In 2017, all stocks but one perfomed well (had a postive return). SEDG and DQ performed the best at nearly 200% returns. TERP was the only stock that had a negative return for the year.

<img width="279" alt="All_Stocks_2017" src="https://user-images.githubusercontent.com/106691255/176981875-69a01292-acc6-4c81-a96b-be3c04ec03cf.png">

### Analysis of Stocks in 2018

In 2018, however, there was a large difference in the performance of the stocks analyzed. ENPH and RUN were the only two stocks that had a positive return for the year. ENPH had performed well the previous year (129.52% return). RUN also had a postive return in 2017, however it was not as drastic (at only 5.55%). All other stocks in 2018 had a negative return.

<img width="278" alt="All_Stocks_2018" src="https://user-images.githubusercontent.com/106691255/176981884-f550a99c-4a46-454f-ab0d-c0f1be0a7ea1.png">

### Challenges and Difficulties Encountered

- Errors with VBA Code
    * Spelling and syntax errors were the most common for me in this analysis. Luckily, Excel prompts you to fix individual lines when an error occurs, making it much easier to fix the mistake.
  
- Limitations of the Data Set:
    * This data is from 2017/2018, making it almost obsolete. The market has changed so much in the last 4/5 years. Even though it is important to know the long term trends when investing, it is even more imperative to see how these companies are performing in current times. A company that did will in 2017 and 2018 could be tanking or even filing for bankrupcy today. On the other hand, a company that was not faring well five years ago may be on the rise today.

## Results

Based off of the results of the analysis of all osberved 'Green Energy' stocks in the years 2017 and 2018, two stood out. ENPH performed the best when considering the performance of both years and RUN was a runner up. However, seeing how poorly all 'Green Energy' stocks performed in 2018, it might be beneficial to look at investing in other types of companies and diversifying your investments. It is clear that the 'Green Energy' stocks can be quite volitle and investing in a range of different types of stocks might be a good idea.

## Summary

### Advantages and Disadvantages of Refactoring Code

Refactoring the coded allowed me to insoect my previous draft in order to find ways to condense, simplify, and improve my code. It also increased the speed at which the sub ran, but not by much. The time it took to re-write the code was much longer than the time it saved in processing. It would have been much more advantageous to refactor the code were I attempting to analyize a much larger data set (an exponentially larger number of tickers to search for). 

### Advantages and Disadvantages of the Original and Refactored VBA Script

The original code was beneficial because it allowed me to think through each component of the analysis piece by piece. It did, however, have some additional, unnecessary information which I was able to condense in the refactoring process.

Again, the refactoring helped the program run a bit quicker, but not overwhelmingly so. Overall, the original draft of code was fine and would have worked well in this situation.

The run times of the original code:

<img width="269" alt="VBA_Challenge_2017_Original" src="https://user-images.githubusercontent.com/106691255/176982246-754b4f31-672d-45a6-bfd3-fa9bcccc238f.png">

<img width="269" alt="VBA_Challenge_2018_Original" src="https://user-images.githubusercontent.com/106691255/176982252-ba24244c-cfa5-4669-ba2c-ed9b5ddc2a24.png">

Run times of the refactored code:

<img width="273" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/106691255/176982208-fd0f14d0-1582-49f2-ad70-43d98cb45930.png">

<img width="275" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/106691255/176982203-9086841a-bb6a-47ed-be08-cecf55e9cbbd.png">


