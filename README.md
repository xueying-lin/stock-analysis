# VBA Challenge-Stock Analysis

## Overview of Project
      Steve recently is graduated with a finance degree, and he is looking for a good performance stock for his parents to invest.
     
### Purpose
      This data analysis is to help Steve to quickly look over the daily volume and yearly return of certain stock for Year 2017 and Year 2018,
      so Steve can select the best performance Stock with his expertise efficiently.
      Once Steve has a larger database, our code should help him to do the same analysis.

## Results
1. Stock Performance Comparison in 2017 & 2018
- All of the stocks performed much better in 2017 than 2018
- As we can see from the analysis table for Stock 2017  And the table for Stock 2018:
      ![Stock performance in 2017](https://raw.githubusercontent.com/xueying-lin/stock-analysis/main/Stock_2017_Performance.PNG)
      ![Stock performance in 2018](https://github.com/xueying-lin/stock-analysis/blob/17c0cbf294c9f497bddde4d89c126d2d45340368/Stock_2018_Performance.PNG)
       
  In 2017, Except for Stock TERP, All other stocks have a positive year return. The Stock DQ, SEDG, ENPH, and FSLG even gained over 100% year return, which is impressive.
  However, in 2018, only stock ENPH and RUN have a positive year return with around 80%, all the other stocks have a negative year return.
  Therefore, the stock performance is much better in 2017 than in 2018
       
- Based on the yearly return in 2017 and 2018, it seems that ENPH is a good and stable stock to invest since it keeps a positive yearly return even when the market situation is bad. But more year data needs to be get to make a suggestion.

2. Necessity of Refactoring code
-Here is the execution time of original code for stock in 2017 and stock in 2018:
       ![Execution time for original code in 2017](https://github.com/xueying-lin/stock-analysis/blob/17c0cbf294c9f497bddde4d89c126d2d45340368/originalcode_2017.PNG)
       ![Execution time for original code in 2018](https://github.com/xueying-lin/stock-analysis/blob/17c0cbf294c9f497bddde4d89c126d2d45340368/originalcode_2018.PNG)
        
As we can see, the execution time is longer than 1s for both worksheets. So the original code may take longer to analyze the thousands of stock data.

- After refactoring the code, the execution time for stock in 2017 and stock in 2018 is as follows:
        ![Execution time after refactorization in 2017](https://github.com/xueying-lin/stock-analysis/blob/17c0cbf294c9f497bddde4d89c126d2d45340368/VBA_Challenge_2017.PNG)
        ![Execution time after refactorization in 2018](https://github.com/xueying-lin/stock-analysis/blob/17c0cbf294c9f497bddde4d89c126d2d45340368/VBA_Challenge_2018.PNG)
        
The execution time is reduced by almost 1s for both worksheets. Therefore, refactoring code is very necessary for future analysis of thousands of stock information.
By refactoring data, we can improve the work efficiency.

## Summary
1. What are the advantages or disadvantages of refactoring code?
- The **advantages** could be:
    - **Time Saver**: reduce the code running time
    - **High Maintanability**: the code is easy to enhance and maintain in the future, so it can handle a large dataset
    - **Clear and Neat**: refactoring code can remove duplicated code, long methods, large classes, etc., and make it hard to read
    - **Fix Bug**: as the code is restructured, the bugs lead by duplicated code and long methods could be removed
       
 - However, there is some **disadvantages** of refactoring code:
    - **Introcude bugs**: since the code is very condensed, we may need to be cautious of code structure. The order of the code may introduce different bugs.
    - **Cost development time**: the programmers need to take time to refactoring code

2. How do these pros and cons apply to refactoring the original VBA script? 
- **Pros:**
    - As discussed in the *Results* section, the execution time is reduced by 1s for both worksheets
    - Avoid nested for loop, which is easy to cause bugs
        - Original code:
``` For i = 0 To 11
              ticker = tickers(i)
              totalVolume = 0
              Sheets(yearValue).Activate
    
               For j = rowStart To rowEnd
      
                             If Cells(j, 1).Value = ticker Then
                                        totalVolume = totalVolume + Cells(j, 8).Value
                             End If
        
          
                             If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
                                       startingPrice = Cells(j, 6).Value
                             End If
       
         
                             If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
                                       endingPrice = Cells(j, 6).Value
                             End If
       
                  Next j
     
                  Worksheets("All Stocks Analysis").Activate
                  Cells(4 + i, 1).Value = ticker
                  Cells(4 + i, 2).Value = totalVolume
                  Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
       
             Next i
 ```
- Refactoring code:
 ```
          For j = 0 To 11
               tickerVolumes(tickerIndex) = 0
          Next j
       
          For i = 2 To RowCount
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
       
                If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                           tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                End If
           
                If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                          tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

                          If Cells(i + 1, 1).Value <> Cells(i - 1, 1).Value Then
                                  tickerIndex = tickerIndex + 1
                         End If
                 End If    
    
             Next i
 ```
-  **Cons:**
   - When refactoring the code, there is an error pop up due to the wrong order of the code as follows:
    >run-time error 9 subscript out of range
   - The wrong code order was:
 ```
          If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                          tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
          End If

          If Cells(i + 1, 1).Value <> Cells(i - 1, 1).Value Then
                           tickerIndex = tickerIndex + 1
          End If
 ```
- Under this code, once the next row cell is not equal to previous row cell, the tickerIndex will increase. Based on our excel data, it will lead to tickerIndex into 12, which is out of range of this variable.
    - To avoid this error, I should the second if condition into the first if condition, making a nested if-then statement.
    - Or, I can specify that Cells(i, 1).Value must equal to current ticker and if the next ticker is different from the previous, we should update the tickerIndex.
        - Like following code:
```
              If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                          tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

                          If Cells(i + 1, 1).Value <> Cells(i - 1, 1).Value Then
                                  tickerIndex = tickerIndex + 1
                         End If
                 End If    
```
            
From this bug I encountered, it is obvious that reconfactoring code could introduce new bugs. And it takes time to develop a correct code.
