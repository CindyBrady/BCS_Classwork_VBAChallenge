
Sub tickerloop()

'Loop through all the sheets.
    For Each ws In Worksheets

        'Set variables
        Dim tickername As String
        Dim tickervolume As Double
        tickervolume = 0
        Dim summary_ticker_row As Integer
        summary_ticker_row = 2
        Dim open_price As Double
        open_price = ws.Cells(2, 3).Value
        Dim close_price As Double
        Dim yearly_change As Double
        Dim percent_change As Double

        'Label the Summary Table headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        'Count the number of rows in the first column
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Loop through the rows by the ticker names
        For i = 2 To lastRow

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
              tickername = ws.Cells(i, 1).Value

              tickervolume = tickervolume + ws.Cells(i, 7).Value

              ws.Range("I" & summary_ticker_row).Value = tickername

              ws.Range("L" & summary_ticker_row).Value = tickervolume

              close_price = ws.Cells(i, 6).Value

              yearly_change = (close_price - open_price)
              
              ws.Range("J" & summary_ticker_row).Value = yearly_change

              'If statement for division by zero error
                If open_price = 0 Then
                    percent_change = 0
                
                Else
                    percent_change = yearly_change / open_price
                
                End If

              ws.Range("K" & summary_ticker_row).Value = percent_change
              ws.Range("K" & summary_ticker_row).NumberFormat = "0.00%"
   
              'Reset the summary row counter, volume and open price
              summary_ticker_row = summary_ticker_row + 1
              tickervolume = 0
              open_price = ws.Cells(i + 1, 3)
            
            Else
              
                tickervolume = tickervolume + ws.Cells(i, 7).Value

            
            End If
        
        Next i

    'Conditional formatting to highlight red and green
    
    'Count number of rows
    lastrow_summary_table = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'Formatting for red or green
        For i = 2 To lastrow_summary_table
            
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 10
            
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            
            End If
        
        Next i

    'Highlight the stock price changes
    'Label the headers

        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

        For i = 2 To lastrow_summary_table
        
            'Get max percent change
            If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrow_summary_table)) Then
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(2, 17).NumberFormat = "0.00%"

            'Get min percent change
            ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrow_summary_table)) Then
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
            
            'Get max volume of trade
            ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrow_summary_table)) Then
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
            
            End If
        
        Next i
    
    Next ws
        
End Sub
