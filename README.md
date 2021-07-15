# VBA of Wall Street

## Overview of Project
This is an Excel Macro designed to analyze the potential benefit of different stocks from 2017-2018, comparing them by daily volume and % of profit. 

## Results
### 2017 
| Ticker     | Total Daily Volume |Return|
| ---      | ---       | ---|
|AY	|136,070,900|	8.9%|
|CSIQ|	310,592,800|	33.1%|
|DQ|	35,796,200|	199.4%|
|ENPH|	221,772,100|	129.5%|
|FSLR|	684,181,400|	101.3%|
|HASI|	80,949,300|	25.8%|
|JKS|	191,632,200|	53.9%|
|RUN|	267,681,300|	5.5%|
|SEDG|	206,885,200|	184.5%|
|SPWR|	782,187,000|	23.1%|
|TERP|	139,402,800|	-7.2%|
|VSLR|	109,487,900|	50.0%|

### 2018 
| Ticker     | Total Daily Volume |Return|
| ---      | ---       | ---|
|AY|	83,079,900|	-7.3%|
|CSIQ|	200,879,900|	-16.3%|
|DQ|	107,873,900|	-62.6%|
|ENPH|	607,473,500|	81.9%|
|FSLR|	478,113,900|	-39.7%|
|HASI|	104,340,600|	-20.7%|
|JKS|	158,309,000|	-60.5%|
|RUN|	502,757,100|	84.0%|
|SEDG|	237,212,300|	-7.8%|
|SPWR|	538,024,300|	-44.6%|
|TERP|	151,434,700|	-5.0%|
|VSLR|	136,539,100|	-3.5%|

We can clearly see that there´s a substantial difference between the performance of the stock market regarding the green energy companies between 2017 and 2018, 2017 was a great year for most of this stocks, the least favored stock dropped by just -7.9%, nothing to worry if you are a long term investor. In the other hand, 2018, was really tough for all the stocks, for the ones that lost value there´s an average drop of -26%, and that should worry you even if you go long term. 

This macro was refactored from a 1st attempt to one that performed better;
In the original Macro we had the following running times for our analysis

![alt text](https://github.com/MauricioIQA/Module-2/blob/main/2017%20Run%20time%20Before.PNG)

![alt text](https://github.com/MauricioIQA/Module-2/blob/main/2018%20Run%20time%20Before.PNG)

And then, when we refactor it

![alt text](https://github.com/MauricioIQA/Module-2/blob/main/VBA_Challenge_2017.PNG)

![alt text](https://github.com/MauricioIQA/Module-2/blob/main/VBA_Challenge_2018.PNG)

The main difference between the 2 codes was that in the first one, we look through all the tickers for each ticker in the list and in the second code we take the information for all the tickers in a simple run.

### First Attempt
``` 
Sub AllStocksAnalysis()
   '1) Format the output sheet on All Stocks Analysis worksheet
   Worksheets("All Stocks Analysis").Activate
   Range("A1").Value = "All Stocks (2018)"
   'Create a header row
   Cells(3, 1).Value = "Ticker"
   Cells(3, 2).Value = "Total Daily Volume"
   Cells(3, 3).Value = "Return"

   '2) Initialize array of all tickers
   Dim tickers(12) As String
   tickers(0) = "AY"
   tickers(1) = "CSIQ"
   tickers(2) = "DQ"
   tickers(3) = "ENPH"
   tickers(4) = "FSLR"
   tickers(5) = "HASI"
   tickers(6) = "JKS"
   tickers(7) = "RUN"
   tickers(8) = "SEDG"
   tickers(9) = "SPWR"
   tickers(10) = "TERP"
   tickers(11) = "VSLR"
   '3a) Initialize variables for starting price and ending price
   Dim startingPrice As Single
   Dim endingPrice As Single
   '3b) Activate data worksheet
   Worksheets("2018").Activate
   '3c) Get the number of rows to loop over
   RowCount = Cells(Rows.Count, "A").End(xlUp).Row

   '4) Loop through tickers
   For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       '5) loop through rows in the data
       Worksheets("2018").Activate
       For j = 2 To RowCount
           '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
           '5b) get starting price for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If

           '5c) get ending price for current ticker
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               endingPrice = Cells(j, 6).Value

           End If
       Next j
       '6) Output data for current ticker
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

   Next i

End Sub
```


###  Refactored
```
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            

            '3d Increase the tickerIndex.
          tickerIndex = tickerIndex + 1
            
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
```

## Summary

What are the advantages or disadvantages of refactoring code?
One of the greatest advantage with refactoring is that for a lot of the problems you encounter there´s a really big chance that others have faced the same problem, therefore, there's a big chance that if you google it you will find the solution or something that will work for you if you make some slight changes, but this is also risky because maybe the code you find works for you but you don’t know how it works and that can be problematic especially if you want to use the code for other applications 

How do these pros and cons apply to refactoring the original VBA script?
In this case what is more evident is that we found a better solution to something we had, that’s also something you can get from refactoring, you can actually find code that helps you online and maybe make your own code from copied pieces of code, refactoring is a way of making your own code.  
