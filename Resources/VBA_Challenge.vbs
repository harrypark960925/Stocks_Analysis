Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    year_value = InputBox("What year would you like to run the analysis on?")

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
    row_count = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
     ticker_index = 0 
        
    '1b) Create three output arrays   
    Dim ticker_volumes as Long 
    Dim ticker_starting_prices as Single 
    Dim ticker_ending_prices as Single 
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero. 
    for i=0 to 11 
        ticker_volumes(i) = 0 
    next i 
        
    ''2b) Loop over all the rows in the spreadsheet. 
    For i = 2 To row_count
    
        '3a) Increase volume for current ticker
    ticker_volumes(ticker_index) = ticker_volumes(ticker_index) + Cells(i,8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i,1).Value = tickers(ticker_index) and Cells(i-1,1).Value <> tickers(ticker_index) Then 
            tickers_starting_prices(ticker_index) = Cells(i,6).Value 
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        If Cells(i,1).Value = tickers(ticker_index) and Cells(i+1,1).Value <> tickers(ticker_index) Then 
            
            tickers_ending_prices(ticker_index) = Cells(i,6).Value 
    
            '3d Increase the tickerIndex. 
            
            ticker_index = ticker_index + 1
        End If
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(i + 4, 1).Value = tickers(i)
        Cells(i + 4, 2).Value = ticker_volumes(i)
        Cells(i + 4, 3).Value = ticker_ending_prices(i) / ticker_starting_prices(i) - 1
        
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
