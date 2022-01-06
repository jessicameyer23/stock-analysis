# Stock Analysis

## Overview of Project: Purpose and Background
The purpose of this project is to analysis many different stocks' volume and returns from 2017 and 2018 to help Steve make a recommendation to his parents.  Steve's parents originally wanted to invest in DQ stock, however, they may change their mind after reviewing DQ's stock performance and related trading volumes compared to other similar stocks.

### Results
## 2017 Results
In 2017, DQ stands out as a strong performer, with a return of 199%, however, the total daily volume was much smaller than the other stocks, which is a red flag.  I would recommend that we analyze the DQ's performance over a longer period of time and continue to compare it to other stocks.  Also, in general, all stocks analyzed had favorable results, but there are large differences in returns among the different companies.

![image_name](https://github.com/jessicameyer23/stock-analysis/blob/main/Resources/2017%20Picture%20of%20Returns%202022-01-06%20075403.png)

## 2018 Results
In 2018, there were only two stocks with postive retuns, ENPH and RUN.  Both of these stocks had favorable returns in 2017, however there was a larger change in the rate of return for RUN, increasing from a return of 5.5% in 2017 to 84% in 2018 with a very large volume of stock traded daily.
In 2018, DQ did not have a favorable return, had a negative return of 62.6%, however the daily trading volumes were much higher, increasing from $35.8M in 2017 to $107.9M in 2018.  
Overall, I would recommend that we analyze another year of data as well as to investigate the causes of the negative returns in 2018 to the industry.  

![image_name](https://github.com/jessicameyer23/stock-analysis/blob/main/Resources/2018%20Picture%20of%20Returns%202022-01-06%20075516.png)


Overall, before moving foward with any investment, I would recommend that we analyze another year of data as well as to investigate the causes of the negative returns in 2018 to the industry.

## Summary

   ## Refactoring Code in General:
1. In general, refactoring code makes it more efficient and can help reduce memory and the time it takes to run code.  It can also make it easier for others to read in the future.  One example of this could be a code that was written to review down to line 57 in an excel document.  However, the excel document may contain more rows over time; therefore, refactoring the code and changing the code to look at the end row in the file will be more efficent and could save error debugging time in the future. 
2. Refactoring in general can also improve the design, and help find bugs.
3. One disadvantage or consideration for refactoring code is that sometimes the person refactoring the code is not the person who originally wrote the code.  The new person refactoring the code has to rely on the original programmer's notes and descriptions. If the person who wrote the original code doesn't leave good explanations or descriptions, there could be a risk that the person refactoring the code could change something that actually doesn't make the code more efficient or could cause a problem down the road, but they would not become aware of that until later.
4. Another disadvantage is that it requires you to retest the functionality and adds more time.

  ## Advantages and Disadvantages of my Original and Refactored VBA script:
1.  For this particular analysis, the original code (which I have included in my script as "Refractored Original" for comparison) looped through all of the data and then returned information only for the specific ticker it was looking for and then continued to repeat that analysis for each ticker.  When I refactored my code (See "Refactored" code), the analysis instead returned information for each ticker as it was looping through the information.  This is more efficient and enabled me to reduce my run time.
2.  Refactoring the code also decreases the likelhood for errors in the coding.  
3.  Refactoring the code however, did increase my (programmer) time spent on refactoring.  Due to my lack of experience, it took a lot of trial and error time, so though it made the actual code more efficient, it took more time on the front end.  

2017 New run time

The new refactored analysis reduced my 2017 run time from .585 seconds (See "All Stocks Analysis" tab for related screenshot on run time in "VBA Challenge xlms") to .445 seconds (see below).  The refactoring reduced my run time by over .1 second, which could be very significant for large amounts of data analysis.

![image_name](https://github.com/jessicameyer23/stock-analysis/blob/main/Resources/2017%202022-01-05%20173432.png)


2018 New Run time

The new refactored analysis reduced my 2018 run time from .57 seconds (See "All Stocks Analysis" tab for related screenshot on run time in "VBA Challenge xlms") to .48 seconds (see below).  The refactoring reduced my run time by almost .1 second, which could be very significant for large amounts of data analysis.

![image_name](https://github.com/jessicameyer23/stock-analysis/blob/main/Resources/2018%202022-01-05%20173305.png)
