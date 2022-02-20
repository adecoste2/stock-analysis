# Stock-analysis Using VBA in Excel

## Overview of Project

### Purpose
The purpose of this project was to analyse green energy stock investments for 2017 and 2018 for the client who is a financial planner to easily uncover the best perfomring green energy stocks for his clients. This process was executed by first writing macro code in VBA to provide the ticker, the total daily volume and the return on each stock. For ease of use and maximum efficiency for the client the VBA code was then refractored. 

### Excel Data
The data in the excel workbook given by the client (green_stocks.xlsx) included two stock tabs, one for 2017 and one for 2018 each with the ticker name, date, the stock openning value, the stock closing value, the adjusted closing value and the volume.   

## Results
To start I copied the code that was needed to create the input box, headers rows for the chart, coded the ticker array, and designated the appropriate worksheet to active. I then removed any unessisary modules in the VBA Project window to streamline efficiency for the client. Below is the code as written in the file.

Sub AllStocksAnalysisRefactored()

    Dim StartTime As Single
    Dim SecondsElapsed As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    StartTime = Timer
    
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
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
   
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         End If

            '3d Increase the tickerIndex.
             If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
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
    MsgBox "This code ran in " & (endTime - StartTime) & " seconds for the year " & (yearValue)

End Sub

## Summary

### General Advantages and Disadvantages of Refractoring Code
The advantages of refractoring code is that it makes it easier to read, for ourselves and others in addition to cutting down the runtime. This intern makes the macro much more efficient. 
Common disadvantages of refractoring code is that it can be time consuming and can be risky when the application is big or the developer does not understand what the original code is trying to accomplish. 


### Advantages and Disadvantages of Refractoring Green Stock Analysis
The advantage of refractoring the Green Stock Analysis VBA code was the decline of the macro runtime and more easily accessable read. The disadvantage of refractoring the Green Stock Analysis VBA code was time - it was an added step that cost time. However, overall refractoring the code is worth the time to make it more accessable and faster for the client. Below are the screenshots with the refractored run time.

![VBA_Challenge_2017](VBA_Challenge_2017.png)
![VBA_Challenge_2018](VBA_Challenge_2018.png)




