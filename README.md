# stock-analysis
# Refactor VBA code and Measuring Code Performance 

## Overview of Project
Our client, is Steve whose parents are interested in investing in green energy stocks.
Steve initally asked for our help in analyzing a handful of green energy stocks, 
but now he wants to expand the dataset to include the entire stock market for the last few years.  

This project included writting a macro in VBA to summarize the performance of stocks as measured by Starting Price, Ending Price and percent return for each stock(by ticker). 
Through the code, this information is summarizes in an Excel worksheet with formatting to highlight positive and negative return. 
The code that was originally written to analyze the performance of a specific stock was refactored, i.e., rewritten for improvement and efficiency, 
so that Steve can easily and quickly analyze the yearly performance of several stocks. 

## Results: Comparison of stock perfomance between 2017 and 2018

Execution times of the original script and refactored script

The execution time of the original All Stock Analysis for 2017 data was .828125 seconds. The execution time was improved with the refactored code which had a run time of .1171875 seconds.

The execution time of the original All Stock Analysis for 2018 data was .875 seconds. The refactored execution time of 1328125 also showed significant improved in run time. 

! [This is a Screen Shot of 2017 Stock Performance Table and Original Script Execution Time] (https://github.com/jbpitman/stock-analysis/blob/main/AllAnalysisOriginal_2017ScreenShot.png)
! [This is a Screen Shot of 2018 Stock Performance Table and Original Script Execution Time] (https://github.com/jbpitman/stock-analysis/blob/main/AllAnalysisOriginal_2018ScreenShot.png)

! [This is a Screen Shot of 2017 Stock Performance Table and Refactored Script Execution Time] (https://github.com/jbpitman/stock-analysis/blob/main/AllAnalysisRefactored_2017ScreenShot.png)

! [This is a Screen Shot of 2017 Stock Performance Table and Refactored Script Execution Time] (https://github.com/jbpitman/stock-analysis/blob/main/AllAnalysisRefactored_2018ScreenShot.png)

### Conclusions
Advantages and Disadvantages of Refactoring code

According to feedback on Stackflow[^1], an obvious reason for refactoring code is that the code runs faster, which can have financial implications (time is money).  
Refactoring often fixes bugs and make the code better by using the latest coding techniques.[^2]   Some coders believe the refactoring code is just part of coding "health", i.e., it shold be done
periodically as part of maintaining the code. 

Although refactoring code should result in faster run times, the benefit to a user may not be obvious. If funding or finding time  to perform refactoring code are issues,
the pros (benefits) of refactoring code greatly exceed the cons. Also, refactoring of a large code could introduce more opportunities to introduce errors. 

How do the pros and cons apply to refactoring the original VBA script 

In the case of the stock analysis, the refactored code did run faster. My first attempt to refactor resulted in additional errors that I did not encounter with the original code.
I was able to go back and correct the error; however, this took more time and effort.  The pop up box with the run time tells the user how long it took to run the program, but aside from this,the average user problably wouldn't be able to tel how the code was improved.  


[^1] https://stackoverflow.com/questions/43983284/what-are-the-advantages-and-disadvantages-of-refactoring-code-smell-in-software

[^2] https://www.quora.com/What-are-the-pros-and-cons-of-refactoring
