Sub Stocks()

'loop through all the sheets
For Each ws In Worksheets

    'set an initial variables
    Dim ticker_name As String
    Dim year_open As Double
    Dim year_close As Double
    Dim stock_vol As Double
    Dim year_change As Double
    Dim year_percentage As Double
    
    'keep track of the location for each stock in the summary table
    Dim summary_table_row As Integer
    summary_table_row = 2
    
    'fill out summary table headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    'identify the last row in the table
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'set the first opening price
    year_open = ws.Range("C2").Value
    
    'loop through all stocks
    For i = 2 To lastrow
        
        'check if we are within the same stock, if we are NOT
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'set stock name
            ticker_name = ws.Cells(i, 1).Value
            
            'identify the closing price
            year_close = ws.Cells(i, 6).Value
            
            'identify year change and change percentage
            year_change = year_close - year_open
            
            'identify change percentage, you can't divide by 0
            If year_close = 0 Then
                year_percentage = 0
            Else
                year_percentage = year_change / year_close
            End If
            
            'add to the stock volume
            stock_vol = stock_vol + ws.Cells(i, 7).Value
        
            'fill in ticker symbol in the summary table
            ws.Range("I" & summary_table_row).Value = ticker_name
            
            'fill in yearly change in the summary table
            ws.Range("J" & summary_table_row).Value = year_change
            ws.Range("J" & summary_table_row).NumberFormat = "0.00"
            
            'fill in percent change in the summary table
            ws.Range("K" & summary_table_row).Value = year_percentage
            ws.Range("K" & summary_table_row).NumberFormat = "0.00%"
            
            'fill in stock volume in the summary table
            ws.Range("L" & summary_table_row).Value = stock_vol
            
            'add a row to the summary table
            summary_table_row = summary_table_row + 1
            
            'reset to the opening price
            year_open = ws.Cells(i + 1, 3).Value
            
            'reset the stock volume
            stock_vol = 0
        
        'If the cell immediately following a row is the same name
        Else
            
            'add to the stock volume
            stock_vol = stock_vol + ws.Cells(i, 7).Value
            
        End If
    
    Next i

    'CONDITIONAL FORMATTING
    'change width of the columns in the summary table
    ws.Range("I:L").EntireColumn.AutoFit

    'identify the last row in the summary table
    lastrow_summary = ws.Cells(Rows.Count, 9).End(xlUp).Row

    'color yearly change based on positive or negative change
    For j = 2 To lastrow_summary
        If ws.Cells(j, 10).Value < 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 3
        Else
            ws.Cells(j, 10).Interior.ColorIndex = 4
        End If
    Next j

    'CHALLENGE
    'create new table headers
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"

    'loop through the summary table
    For y = 2 To lastrow_summary
    
        'identify the maximum percent change and fill in new table
        If ws.Cells(y, 11).Value = WorksheetFunction.Max(ws.Range("K2:k" & lastrow_summary)) Then
            ws.Range("O2").Value = ws.Cells(y, 9).Value
            ws.Range("P2").Value = ws.Cells(y, 11).Value
            ws.Range("P2").NumberFormat = "0.00%"
        
        'identify the minimum percent change and fill in new table
        ElseIf ws.Cells(y, 11).Value = WorksheetFunction.Min(ws.Range("K2:K" & lastrow_summary)) Then
            ws.Range("O3").Value = ws.Cells(y, 9).Value
            ws.Range("P3").Value = ws.Cells(y, 11).Value
            ws.Range("P3").NumberFormat = "0.00%"
            
        'identify the maximum total stock volume and fill in new table
        ElseIf ws.Cells(y, 12).Value = WorksheetFunction.Max(ws.Range("L2:L" & lastrow_summary)) Then
            ws.Range("O4").Value = ws.Cells(y, 9).Value
            ws.Range("P4").Value = ws.Cells(y, 12).Value
        End If
        
    Next y

    'change width of the columns in the new table
    ws.Range("N:P").EntireColumn.AutoFit

Next ws

End Sub