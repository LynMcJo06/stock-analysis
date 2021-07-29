# Stock Analysis
## Introduction
In this module, our client Steve is a recent graduate with a Finance degree.  Steve's parents are interested in green energy and want to start investing.  One company caught their attention:  DAQO New Energy Corp (DQ).  Steve's parents want to invest only in this stock, which has Steve concerned.  Steve has asked for our help to determine if DQ is a good investment.  
### Lesson 1
We are going to start by analyzing a variety of clean energy stocks just in case DAQO turns out negatively.  We converted our Excel spreadsheet to a macro-enabled sheet and learned several new terms.  To ensure that VBA was working properly, we created a simply macro that printed "Hello World!"  Woohoo!  I have created my first piece of code.  
### Lesson 2
In this lesson, we started analyzing the stock data.  We first added a new worksheet titled "DQ Analysis", to display our output analysis.  our macro started simple and gradually grew as our knowledge of the VBA coding language grew.  Steve's parents are convinced that a stock is a good investment if it is traded frequently, so we determined the total daily volume of DQ stock for 2018.   In 2018, DQ's Total Daily Volume was 107,873,900.  We also calculated DQ's 2018 yearly return.  This portion required a nested for loop incorporating logical operators and conditional statements.  We researched how to create a portion of our code via the Internet.  The results are documented in the code block.  After running our DQ Analysis code block, DQ's early return was -0.626, which equates to a loss of approximately 63%.  We are concerned about Steve's parents investing in DQ.  
### Lesson 3
For this lesson, we analyzed all stocks to get their total daily volume and 2018 yearly return.  We were able to reuse a lot of our code from the previous lesson by making minor changes.  After we ran this code block, only two of the eleven stocks had a positive yearly return:  ENPH and RUN, 82% and 84%, respectively.  
### Lesson 4
In this lesson, we formatted the All Stocks Analysis sheet to make it readable.  We created a new macro and added our own formatting as well as the coursework.  
### Lesson 5
In Lesson 5, we learned how to make the worksheets interactive.  We first created buttons in the All Stocks Analysis and DQ Analysis sheets.  The buttons cleared the sheet and ran the analysis.  We then created an Input Box so the analysis could be performed on stocks for any year.  I tried it using both 2017 and 2018 and it worked wonderfully.  For the final exercise, we created code to calculate the elasped time it takes for the All Stocks Analysis code to run.  For the 2018 datasheet, the runtime was 1.17 seconds.  The runtime for the 2017 datasheet was 1.04 seconds.  It should be noted that the formatting code was placed in a separate subroutine, which had to be executed after the initial analysis.  The formatting time was not included in the runtime for either dataset.  
# MODULE 2 CHALLENGE
## Purpose
The purpose of this challenge is to determine if the coursework code can be refactored for a much larger dataset such as the entire stock market. The goal was to loop through the data one time and collect the same information as the coursework.  
## Results
I have to admit, it took a lot of debugging and rewriting to get the refactored code to work.  I initially made several errors, but through perserverance and several weeks of code writting, I was able to accomplish the challenge goals.  My refactored code did collect the same information and at a faster runtime.  For the challenge, I added the output sheet formatting within the runtime variables, which makes the final runtimes even faster.  For the coursework, the output sheet formatting was completed in a separate subroutine; therefore, I had to create an additional button.  

For 2017, the refactored code ran in 0.17 seconds.  
![VBA_Challenge_2017](https://user-images.githubusercontent.com/84352487/127578548-f429aeb5-5111-4057-9bab-7942f306982d.png)

For 2018, the refactored code ran in 0.16 seconds.  
![VBA_Challenge_2018](https://user-images.githubusercontent.com/84352487/127578654-be33036c-44f1-4371-bfe6-990721542399.png)

## Summary
Refactoring the code, even though it took me a long time, has the following advantages:
  - It can be used in future analysis, with several different types of stocks, and for different years
  - The code ususally runs faster--the code is more efficient
  - Refactoring is great practice, especially for novice programmers
  - The results are predictable, which is an easy way to confirm that the refactored code is running properly
  - The refactored code is easier to understand
  - It can support future advances in software development

Some disadvantages of refactoring may include:
  - It can be time consuming
  - The coder may get lost or confused
  - You may not have a test module to reference

The coursework (original) code was simplier and faster to write even though it took slightly longer to execute the end results.  I found the refactoring portion extremely challenging as I did not understand arrays or loops.  This was an excellent exercise and I had to conduct a lot of research, trial and error, and rewriting.  
