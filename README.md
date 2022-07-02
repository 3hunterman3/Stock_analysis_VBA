# VBA_Challenge

Overview of Project:
- The objective was to be able to take the data given to us on various stocks from 2017 and 2018 and see the stocks individual performence. Taking the data we had to see if the company is worth investing and to show each individual companys returns. Also the purpose of this was to be able to increase my understanding macros and coding using VBA syntax and being able to manipulate the code to get the results I need.

Results:
- When comparing the stock performance from 2017 and 2018 we can see very different results:
![total data 2017](/total_2017.png)
![total data 2018](/total_2018.png)
- In the pictures we can see the data for 2017 has more positive returns and for 2018 more negative returns.
- we can look at the exectuion time of our VBA code and it shows that our 2017 data exectued the data in 0.125 seconds compared to out 2018 data which was 0.1289 seconds. We can also compare this execution to our original script compared to the refactored script when we compare the times we get the following times:![original](/original.png)
- we can see that our new refarctored code runs a lot more quickly than that of our original code.

**The code of our refactored VBA code:**
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
    Dim tickers(11) As String
    
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
For j = 2 To RowCount
        
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(j - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row‚Äôs ticker doesn‚Äôt match, increase the tickerIndex.
        'If  Then
        If Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
            tickerIndex = tickerIndex + 1
        End If
        
Next j
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
    
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i) - 1)
    Next i
    
    'Formatting

    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For j = dataRowStart To dataRowEnd
        
        If Cells(j, 3) > 0 Then
            
            Cells(j, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(j, 3).Interior.Color = vbRed
            
        End If
        
Next j
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

**The orignal VBA code:**
Sub AllStocksAnalysis()
    yearValue = InputBox("what year would you like to run the analysis on?")
    Dim startTime As Single
    Dim endTime As Single
    
        startTime = Timer
'1) Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
        Range("A1").Value = "All Stocks (" + yearValue + ")"
        Sheets(yearValue).Activate
        
        'create a header row
        Cells(3, 1).Value = "Year"
        Cells(3, 2).Value = "Total Daily Volume"
        Cells(3, 3).Value = "return"
        
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

            'Color the cell green
            Cells(i, 3).Interior.Color = vbGreen

        ElseIf Cells(i, 3) < 0 Then

            'Color the cell red
            Cells(i, 3).Interior.Color = vbRed

        Else

            'Clear the cell color
            Cells(i, 3).Interior.Color = xlNone

        End If
    
    Next i
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & "seconds for the year " & (yearValue)
    

End Sub

Summary:
What are the advantages or disadvantages of refactoring code?
- The advantages of the refactoring of the code allows for more quick execution and takes up less storage. Also the advantages are that we can have code that is easier to read. Readability is extremly important as it allows for more easier debugging when the code does not work and will allow future individuals to see my code and be able to understan what is going on more clearly. I personally dont see any disadvantages with refactoring the code.
How do these pros and cons apply to refactoring the original VBA script?
- The biggest pro in refactoring is that it decreases the executution time tremendously. 
