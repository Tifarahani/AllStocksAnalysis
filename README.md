## Challange 2-Bootcamp


## Overview of Project
>The purpose of this project was to refactor a Stock Market Dataset with VBA and demonstrate "Total Daily Volume" and "Return" in the year 2017 and 2018 using 
> knowlde of Arrays,For loops, Formatting Cells to know if the Market is good to invest in or not.

## Results
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
    For i = 0 To 11
    tickerIndex = tickers(i)

    '1b) Create three output arrays
    Dim tickerVolumes As Long
    Dim tickerStartingPrices As Single
    Dim tickerEndingPrices As Single
    
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    
    ' Activating the worksheet and assigning zero to ticker volumes
    Worksheets(yearValue).Activate
    tickerVolumes = 0
        
    '2b) Loop over all the rows in the spreadsheet.
    
    Worksheets(yearValue).Activate
    For j = 2 To RowCount
    
    '3a) Increase volume for current ticker

    ' If the next row’s ticker match, increase the tickerVolumes.
    
    If Cells(j, 1).Value = tickerIndex Then
           tickerVolumes = tickerVolumes + Cells(j, 8).Value
    End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
    
    If Cells(j - 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then
        tickerStartingPrices = Cells(j, 6).Value
    End If
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rows ticker doesn't match, increase the tickerIndex.
        'If  Then
        
    If Cells(j + 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then
       tickerEndingPrices = Cells(j, 6).Value
     
        
            '3d Increase the tickerIndex.
                       
    End If
         
        'End If
    
     Next j
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
    Worksheets("All Stocks Analysis").Activate
    
    Cells(4 + i, 1).Value = tickerIndex
    Cells(4 + i, 2).Value = tickerVolumes
    Cells(4 + i, 3).Value = tickerEndingPrices / tickerStartingPrices - 1
        
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

 ---
 ### VBA Analysis 2017
![All Stocks(2017)](https://github.com/Tifarahani/Challange-2-Bootcamp/blob/main/Resources/All%20Stocks(2017).png)
![All Stocks-timer(2017)](https://github.com/Tifarahani/Challange-2-Bootcamp/blob/main/Resources/All%20Stocks-timer(2017).png)

---
 ### VBA Analysis 2018
![All Stocks(2018)](https://github.com/Tifarahani/Challange-2-Bootcamp/blob/main/Resources/All%20Stocks(2018).png)
![All Stocks-timer(2018)](https://github.com/Tifarahani/Challange-2-Bootcamp/blob/main/Resources/All%20Stocks-timer(2018).png)
 
 ---
##  Summary:
>
>
>
>
### What are the advantages or disadvantages of refactoring code?
- [x] Advantages:
* After refactoring, the code is fresher, easier to understand or read, less complex and easier to maintain
* Logical errors easily appear in well structure code that contains nested conditionals and loops.
* It prevents many future defects. 
* Code Size is reduced. 
* Confused coding is properly restructured.
* Will lead to more unit testable code 
- [x] Disadvantages:
- Time consuming
- A long procedure may contain the same line of code in several locations, you can change the logic to eliminate the duplicate lines.
-A logical structure may be duplicated in two or more procedures (possibly via copy & paste coding). When detected, this logic is best moved to a new function and called from the other functions.
-Refactoring process can affect the testing outcomes.

### How do these pros and cons apply to refactoring the original VBA script
- [x] Pros: Refactoring leads to better quality code
- [x] Cons: We have to retest lots of functionality and it takes a lot of time

#### Conclusion: 
Refactoring improves the design of software, makes software easier to understand, helps us find bugs and also helps in executing the program faster.


