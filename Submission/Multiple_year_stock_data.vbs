Sub stock_analysis():
    
    'Declare worksheets
    Dim ws As Worksheet
    For Each ws In Worksheets

        'Declaring all the variables
        Dim stock As String
        Dim next_stock As String
        Dim stock_volume As Double
        Dim quarterly_change As Double
        Dim percent_change As Double
        Dim i As Long
        
        'Keep track of the location for each stock name in the stock summary table
        Dim stock_summary_row As Long
        stock_summary_row = 2
        
        'Declaring total stock volume and Assigning the initial value
        Dim total_stock_volume As Double
        total_stock_volume = 0
        
        'Count total number of rows in the first column.
        Dim last_row As Long
        last_row = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        'Label the stock summary table headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
    
        'Declare and Set initial open price value
        Dim open_price As Double
        open_price = ws.Cells(2, 3).Value
        
        'Loop through each ticker until there is a mismatch
        For i = 2 To last_row
            stock = ws.Cells(i, 1).Value
            next_stock = ws.Cells(i + 1, 1).Value
            stock_volume = ws.Cells(i, 7).Value
            
            If stock <> next_stock Then
                'ticker_name = ws.Cells(i, 1).Value
                'ws.Range("I" & stock_summary_row) = stock
                total_stock_volume = total_stock_volume + stock_volume
                ws.Range("L" & stock_summary_row) = total_stock_volume
                
                'Set close price value
                Dim close_price As Double
                close_price = ws.Cells(i, 6).Value
            
                'Calculate quarterly_change
                
                quarterly_change = close_price - open_price
                ws.Range("J" & stock_summary_row) = quarterly_change
                'Conditional formatting that will highlight positive change in green and negative change in red
                If (quarterly_change > 0) Then
                    ws.Cells(stock_summary_row, 10).Interior.ColorIndex = 4
                ElseIf (quarterly_change < 0) Then
                    ws.Cells(stock_summary_row, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(stock_summary_row, 10).Interior.ColorIndex = 2
                End If
                
                'Calculate percent change

                If open_price = 0 Then
                    percent_change = 0
                Else
                    percent_change = quarterly_change / open_price
                End If
          
                ws.Range("K" & stock_summary_row) = FormatPercent(percent_change)

                'Reset the row counter to next row
                stock_summary_row = stock_summary_row + 1
                'Reset the total_stock_volume to 0
                total_stock_volume = 0
                'Reset the opening price of the next stock
                open_price = ws.Cells(i + 1, 3).Value
            Else
                total_stock_volume = total_stock_volume + stock_volume
            End If
            
        Next i

    
        'Second loop for second stock summary table
         Dim max_price As Double
         Dim min_price As Double
         Dim max_volume As Double
         Dim max_price_stock As String
         Dim min_price_stock As String
         Dim max_volume_stock As String
         Dim j As Long
         Dim last_row_summary_table As Long
         
        'initialize to first row of the stock summary table for comparison
        max_price = ws.Cells(2, 11).Value
        min_price = ws.Cells(2, 11).Value
        max_volume = ws.Cells(2, 12).Value
        max_price_stock = ws.Cells(2, 9).Value
        min_price_stock = ws.Cells(2, 9).Value
        max_volume_stock = ws.Cells(2, 9).Value
        
        last_row_summary_table = ws.Cells(Rows.Count, "I").End(xlUp).Row
        
        For j = 2 To last_row_summary_table
            ' Compare current row to the inits (first row)
            If (ws.Cells(j, 11).Value > max_price) Then
                'We have a new Max Percent Change!
                max_price = ws.Cells(j, 11).Value
                max_price_stock = ws.Cells(j, 9).Value
            End If
            
            If (ws.Cells(j, 11).Value < min_price) Then
                'We have a new Min Percent Change!
                min_price = ws.Cells(j, 11).Value
                min_price_stock = ws.Cells(j, 9).Value
            End If
            
            If (ws.Cells(j, 12).Value > max_volume) Then
                ' We have a new Max Volume!
                max_volume = ws.Cells(j, 12).Value
                max_volume_stock = ws.Cells(j, 9).Value
            End If
        Next j
        
        'Label the summary table headers
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'write out to excel
        ws.Range("P2").Value = max_price_stock
        ws.Range("P3").Value = min_price_stock
        ws.Range("P4").Value = max_volume_stock
        ws.Range("Q2").Value = FormatPercent(max_price)
        ws.Range("Q3").Value = FormatPercent(min_price)
        ws.Range("Q4").Value = max_volume

     
    Next ws
    
End Sub
