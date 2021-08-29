# Module 2 Challenge - VBA_Challenge 
#
## Overview of the project
###The client is trying to evaluate the stocks to invest in using two metrics: a) Total daily volume and b) returns. While there is an existing code that automates the calculations needed, it is unoptimized. The goal of the project is refactor the code to speed up the analyses and summarize the data.
#
##Results
### *Stock analyses* - ENPH and RUN were the only two stocks that consistently provided positive returns for both 2017 and 2018. The rest of the stocks (excluding TERP) while were positive for 2017, were negative for 2018. The returns were different for the two years, despite the volumes being similar.
### *Refactor code results* - The Original code took *0.72s* to run, while the refactored code *0.05s* to run. 
#
##Summary 1
###a) Advantages of the refactoring code - Refactoring code is essentially optimization of existing potentially POC code. It enables optimization of memory usage, as well as performance given the knowledge of input and outputs.
###b) Disadvantages of refactoring code - Refactoring code (as I currently understand it), is meant for optimization of existing POC code. Refactoring prior to understanding the input/output could lead to lots of re-coding.
##Summary 2
###Given the initial code was available, along with results from that code, refactoring made sense. The original code was inefficient due to looping the entire table for each stock, instead of looping the table once and synthesizing the information. Additionally, the original code was flipping b/w worksheets which reduces performance.


![](image.png)
   