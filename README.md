# Stock Analysis

## Overview of Project: Purpose and Background
The purpose of this project is to analysis many different stocks' volume and returns from 2017 and 2018 to help Steve make a recommendation to his parents.  Steve's parents originally wanted to invest in DQ stock, however, they may change their mind after reviewing DQ's stock performance and related trading volumes compared to other similar stocks.

### Results
## 2017 Results

![image_name](https://github.com/jessicameyer23/stock-analysis/blob/main/Resources/2017%20Picture%20of%20Returns%202022-01-06%20075403.png)

## 2018 Results
![image_name](https://github.com/jessicameyer23/stock-analysis/blob/main/Resources/2018%20Picture%20of%20Returns%202022-01-06%20075516.png)


## Summary

   ## Refactoring Code in General:
1. In general, refactoring code makes it more efficient and can help reduce memory and the time it takes to run code.  It can also make it easier for others to read in the future.  One example of this could be a code that was written to review down to line 57 in an excel document.  However, the excel document may contain more rows over time; therefore, refactoring the code and changing the code to look at the end row in the file will be more efficent and could save error debugging time in the future. 
2. One disadvantage or consideration for refactoring code is that sometimes the person refactoring the code is not the person who originally wrote the code.  The new person refactoring the code has to rely on the original programmer's notes and descriptions. If the person who wrote the original code doesn't leave good explanations or descriptions, there could be a risk that the person refactoring the code could change something that actually doesn't make the code more efficient or could cause a problem down the road, but they would not become aware of that until later.
3.
  ## Advantages and Disadvantages of my Original and Refactored VBA script:
1.  For this particular analysis, I the original code (which I have incluced in my script as "Refractored Original" for comparison) looped through all of the data and then returned information only for the specific ticker it was looking for and then continued to repeat that analysis for each ticker.  When I refactored my code (See "Refactored" code), the analysis instead returned information for each ticker as it was looping through the information.  This is more efficient and enabled me to reduce my run time.
2.  Refactoring the code also decreases the likelhood for errors in the coding.  
3.
2017 New run time
The new refactored analysis reduced my 2017 run time from .585 seconds (See "All Stocks Analysis" tab for related screenshot on run time in "VBA Challenge xlms") to .445 seconds (see below).  The refactoring reduced my run time by over .1 second, which could be very significant for large amounts of data analysis.
![image_name](https://github.com/jessicameyer23/stock-analysis/blob/main/Resources/2017%202022-01-05%20173432.png)


2018 New Run time
The new refactored analysis reduced my 2018 run time from .57 seconds (See "All Stocks Analysis" tab for related screenshot on run time in "VBA Challenge xlms") to .45 seconds (see below).  The refactoring reduced my run time by over .1 second, which could be very significant for large amounts of data analysis.
![image_name](https://github.com/jessicameyer23/stock-analysis/blob/main/Resources/2018%202022-01-05%20173305.png)
