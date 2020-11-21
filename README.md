# Stock-Analysis

## Purpose
   The purpose of the Stock Analysis was to compare total daily volume and return percentage for the years 2017 and 2018.


## Analysis 2017
![2017 Stock Analysis](https://github.com/AmirO8/Stock-Analysis/blob/main/Week%202%20Challenge/Resources/2017%20Stock%20Analysis.png)

There are several points we can look at for our analysis the company that had the highest total daily volume, lowest total daily volume, highest return, and lowest return.
In 2017 the lowest total daily volume was DQ with 35,796,200. While DQ had the lowest total daily volume it also had the highest return percentage at 199.4%
The highest total daily volume was SPWR at 782,187,000. TERP had the lowest return at -7.2%. They were the only company to have a negative return in 2017.

## Analysis 2018
![2018 Stock Analysis](https://github.com/AmirO8/Stock-Analysis/blob/main/Week%202%20Challenge/Resources/2018%20Stock%20Analysis.png)

When we take a look at the year 2018 there are more companies in with negative returns compared to the year 2017. ENPH and RUN were the only companies that had a positive return. RUN having the highest at 84% and DQ having the lowest at -62.6%. The reason for this could be that DQ had such a high return at the end of 2017 that they opened very high and could not duplicate that in 2018.

## Why Refactor Code
In VBA refactoring the code makes it easier to read and it executes faster. One step I refactored was referring to my array as a whole Dim tickerVolume(12) As Long is an example. I was able to do this by reading this [Array](https://docs.microsoft.com/en-us/office/vba/language/concepts/getting-started/using-arrays) document. 
I was getting errors overflow, sub or function not defined, and subscript out of range. Since I refractored my code I was able to comment out the errors. Refactoring code is not easy and it is time consuming it took me several hours just to refactor certain blocks of code specifically the "For" loops.

When my code finaly ran it ran quickly as shown below for both years.

![2017 code speed](https://github.com/AmirO8/Stock-Analysis/blob/main/Week%202%20Challenge/Resources/VBA_Challenge_2017.png)
![2018 code speed](https://github.com/AmirO8/Stock-Analysis/blob/main/Week%202%20Challenge/Resources/VBA_Challenge_2018.png)

