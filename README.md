# VBA of Wall Street
MS Excel with VBA codes to run stock analysis
## Overview of Project
This project has the purpose to analyze some list of stocks performance, over the year of 2017 and 2018. Using the database to run a high performance macro, written on MS Visual Basic. This scrip is able to run the Return of End stock price and Starting price over the requested year. As well as Total volume. 
## Results
This Analysis shows the stock **End** price divided by **Start** price -1 that give us the *Stock Return* , *positive* (green) or *negative*(red).
### Code
Sub AllStocksAnalysisRefactored()   
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    
Worksheets("All Stocks Analysis").Activate
    
    'Title Analysis
    
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
    
    'Count the number of rows to loop over
    
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker index to reference proper ticker in the arrays.
    
    Dim tickerIndex As Integer
    
    'Initiate tickerIndex at zero.
    
    tickerIndex = 0
    
    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    '2a) Create for loop to analyze each ticker in the array.
    
    For tickerIndex = 0 To 11
    
    'Initiate each ticker's volume at zero.
    
    tickerVolumes(tickerIndex) = 0
    
    'Activate data worksheet
    
    Worksheets(yearValue).Activate
        
        '2b) Loop over all the rows in the spreadsheet.
        
        For i = 2 To RowCount
        
            '3a) Increase volume for current ticker.
            
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
        
            '3b) Check if the current row is the first row with the current ticker.
                    
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            
                'if it is the first row for current ticker, set starting price.
                
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            'End If
            End If
            
            
        '3c) Check if the current row is the last row with the current ticker.

            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
                'if it is the last row for current ticker, set ending price.
                
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            'End if
            End If
            
        '3d) Check if the current row is the last row with the current ticker.
        
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                
                'if it is, increase tickerIndex to move on to next ticker in array.
                
                tickerIndex = tickerIndex + 1
            
            'End If
            End If
    
        Next i
        
    Next tickerIndex
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
    For i = 0 To 11
        
        'Activate Output Worksheet
        
        Worksheets("All Stocks Analysis").Activate
        
        'Ticker Row Label
        
        Cells(4 + i, 1).Value = tickers(i)
        
        'Sum of Volume
        
        Cells(4 + i, 2).Value = tickerVolumes(i)
        
        'ReturnValue
        
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

### 2017 Stocks
![All_Stocks_2017](https://user-images.githubusercontent.com/92833805/140678309-1741aee4-7f1d-411f-ae9e-6e0acb6868dc.png)
###
Great performance on these stocks, mostly green that means postive returns on year of 2017.
### 2018 Stocks
![All_Stocks_2018](https://user-images.githubusercontent.com/92833805/140678417-378b79ff-a866-42c8-8d18-26b34d41e05e.png)
###
Contrary of 2017, the year of 2018 show many negative returns in red meaning a bad performance for selected stocks.
### Execution Times
Running the code successfully, the execution to run the macro for the year of **2017** was 0.08 seconds, while **2018** took 0.09 second. Really fast and light script.
### Execution time for 2017
![VBA_Challenge_2017](https://user-images.githubusercontent.com/92833805/140678369-91eedd6c-cd6a-4e3a-9bcb-6c04780b2ccb.png)
### Execution time for 2018 
![All_Stocks_2018](https://user-images.githubusercontent.com/92833805/140678491-40629f97-d317-4ae5-a49c-def3d72638d9.png)

## Summary
### Advantages of refactoring a code
When you have a good code written, with commented explaining what the line is doing it is very satisfactory and time saving.
### Disadvantage of refactoring a code
Can be very tricky to the reader to understand what of each command are doing, sometime can be very exhaustive to debug the whole script. Although that also a good opportunity to learn different ways to execute.
