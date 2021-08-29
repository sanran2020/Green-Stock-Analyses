# Module 2 Challenge - VBA_Challenge 
#
## Overview of the project
#### The client is trying to evaluate the stocks to invest in using two metrics: a) Total daily volume and b) returns. While there is an existing code that automates the calculations needed, it is unoptimized. The goal of the project is refactor the code to speed up the analyses and summarize the data.
#
## Results
#### *Stock analyses* - ENPH and RUN were the only two stocks that consistently provided positive returns for both 2017 and 2018. The rest of the stocks (excluding TERP) while were positive for 2017, were negative for 2018. The returns were different for the two years, despite the volumes being similar.

<img width="206" alt="2017 Stock Performance" src="https://user-images.githubusercontent.com/89116627/131235560-40e68d7a-a23c-494a-ad2f-9b5ca6a915a4.PNG"> <img width="198" alt="2018 Stock Performance" src="https://user-images.githubusercontent.com/89116627/131235577-6bec093d-cb32-4af6-b826-64653a32c0aa.PNG">


#### *Refactor code results* - The Original code took *0.72s* to run, while the Refactored code *0.05s* to run. 
######
###### Original Code <img width="226" alt="Original VBA_2017" src="https://user-images.githubusercontent.com/89116627/131235630-94d32045-84ff-401e-a371-7a8eb5cf989c.PNG"> Refactored Code <img width="239" alt="refactored time 2017" src="https://user-images.githubusercontent.com/89116627/131235634-5fd38289-31d3-4701-acbb-a140d2f637f7.PNG">
######
#### *Refactor code improvements* - Two key improvements - (a) Not looping over the entire table for each stock (b) Flipping b/w sheets for input and output updates
######
##### (a) Unoptimized code loop 

           '5c) get ending price for current ticker
              If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                  endingPrice = Cells(j, 6).Value
              End If
       Next j
######
##### (a) Refactored code loop

           '3c) check if the current row is the last row with the selected ticker
            'If the next row’s ticker doesn’t match, increase the tickerIndex.
               If (Cells(i + 1, 1) <> tickers(tickerIndex) And Cells(i, 1) = tickers(tickerIndex)) Then
                  tickerEndrow(tickerIndex) = Cells(i, 6)
                  tickerEnddate(tickerIndex) = Cells(i, 2)

            '3d Increase the tickerIndex.
                  tickerIndex = tickerIndex + 1
                  t = 0
            
               End If
######
##### (b) Unoptimized code flip sheets 

          '4) Loop through tickers
           For i = 0 To 11
              ticker = tickers(i)
              totalVolume = 0
           '5) loop through rows in the data
               Worksheets("2018").Activate
           '6) Output data for current ticker
               Worksheets("All Stocks Analysis").Activate 
##### (b) Refactored code flip sheets

            'Activate data worksheet
               Worksheets(yearValue).Activate
               
            '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
               Worksheets("All Stocks Analysis").Activate
               For i = 0 To 11
                  Cells(4 + i, 1) = tickers(i)
                  Cells(4 + i, 2) = tickerVolume(i)
                  Cells(4 + i, 3) = ((tickerEndrow(i) - tickerStrow(i)) / tickerStrow(i))
                Next i

#
## Summary 1
a) Advantages of the refactoring code - Refactoring code is essentially optimization of existing potentially POC code. It enables optimization of memory usage, as well as performance given the knowledge of input and outputs.

b) Disadvantages of refactoring code - Refactoring code (as I currently understand it), is meant for optimization of existing POC code. Refactoring prior to understanding the input/output could lead to lots of re-coding.
## Summary 2
Given the initial code was available, along with results from that code, refactoring made sense. The original code was inefficient due to looping the entire table for each stock, instead of looping the table once and synthesizing the information. Additionally, the original code was flipping b/w worksheets which reduces performance.


   
