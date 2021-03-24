Sub Wall_Street():

Dim ws As Worksheet

For Each ws In Worksheets

    ' Set an initial variable for holding the ticker
    Dim stock_ticker As String
    
    'Set an initial variable for holding the total volume
    Dim vol As Double
    vol = 0
    
    ' Set initial variables for holding data of each stock
    Dim opening_price As Double
    Dim closing_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim yearly_change_max As Double
    Dim yearly_change_min As Double
    Dim vol_max As Double
        
    ' Keep track of the location for each stock in the summary
    Dim summary_table_row As Long
    summary_table_row = 2
    
    Dim max_row As Long
    Dim min_row As Long
    Dim vol_row As Long
    
    ' Set i as a Long Variable
    Dim i As Long
    
    ' Set a variable to track the first and last element of each ticker
    Dim counter As Long
    counter = 2
        
    ' Set Last Row variable
    Dim LastRow As Long
    LastRow = 0

        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
                        
       ' Loop through all tickers
        For i = 2 To LastRow
        
            ' Check if we are still within the same ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                ' Set ticker
                stock_ticker = ws.Cells(i, 1).Value
            
                ' Find opening price
                opening_price = ws.Cells(counter, 3).Value
                
                ' Find closing price
                closing_price = ws.Cells(i, 6).Value
                
                ' Calculate yearly change
                yearly_change = closing_price - opening_price
                
                ' Calculate percent change
                
                If opening_price = 0 Then
                
                    percent_change = 0
                    
                Else
                                
                percent_change = ((closing_price - opening_price) / opening_price)
                
                End If
                
                ' Add to the volume
                vol = vol + ws.Cells(i, 7).Value
            
                ' Print the stock ticker
                ws.Cells(summary_table_row, 10).Value = ws.Cells(i, 1).Value
            
                ' Print the yearly change
                ws.Cells(summary_table_row, 11).Value = yearly_change
                
                ' Print the percent change
                ws.Cells(summary_table_row, 12).Value = percent_change
                
                ' Print the volume
                ws.Cells(summary_table_row, 13).Value = vol
            
                ' Print opening price
                ws.Cells(summary_table_row, 15).Value = opening_price
                
                ' Print closing price
                ws.Cells(summary_table_row, 16).Value = closing_price
                
                ' Add one to the summary table row
                summary_table_row = summary_table_row + 1
            
                ' Reset the volume total
                vol = 0
                
                ' Set initial range roe
                counter = i + 1
            
            ' If the cell immediately following a row is the same ticker
            Else
        
                ' Add to the volume total
                vol = vol + Cells(i, 7).Value
                                       
            End If
            
        Next i
             
        ' Headers
        ws.Cells(1, 10).Value = "Ticker"
        ws.Cells(1, 11).Value = "Yearly Change"
        ws.Cells(1, 12).Value = "Percent Change"
        ws.Cells(1, 13).Value = "Total Stock Volume"
        ws.Cells(1, 15).Value = "Opening Price"
        ws.Cells(1, 16).Value = "Closing Price"
        ws.Cells(1, 19).Value = "Ticker"
        ws.Cells(1, 20).Value = "Value"
        
        ' Greatest indicators
        
        ' Max % increase
        yearly_change_max = WorksheetFunction.Max(ws.Range("L2:L" & summary_table_row))
        max_row = WorksheetFunction.Match(yearly_change_max, ws.Range("L2:L" & summary_table_row), 0)
        ws.Cells(3, 18).Value = "Greatest % Increase"
        ws.Cells(3, 19).Value = ws.Cells(max_row, 10).Value
        ws.Cells(3, 20).Value = yearly_change_max
        ws.Cells(3, 20).Style = "Percent"
        
        ' Min % decrease
        yearly_change_min = WorksheetFunction.Min(ws.Range("L2:L" & summary_table_row))
        min_row = WorksheetFunction.Match(yearly_change_min, ws.Range("L2:L" & summary_table_row), 0)
        ws.Cells(4, 18).Value = "Greatest % Decrease"
        ws.Cells(4, 19).Value = ws.Cells(min_row, 10).Value
        ws.Cells(4, 20).Value = yearly_change_min
        ws.Cells(4, 20).Style = "Percent"
        
       ' Max total vol
        vol_max = WorksheetFunction.Max(ws.Range("M2:M" & summary_table_row))
        vol_row = WorksheetFunction.Match(vol_max, ws.Range("M2:M" & summary_table_row), 0)
        ws.Cells(5, 18).Value = "Greatest Total Volume"
        ws.Cells(5, 19).Value = ws.Cells(vol_row, 10).Value
        ws.Cells(5, 20).Value = vol_max
        
        ' Formatter
        
            ws.Columns("J:T").AutoFit
            
            For i = 2 To summary_table_row
            
                ws.Cells(i, 12).Style = "Percent"
                
                If ws.Cells(i, 11).Value < 0 Then
                
                    ws.Cells(i, 11).Interior.ColorIndex = 3
                    
                Else
                
                    ws.Cells(i, 11).Interior.ColorIndex = 4
                    
                End If
                
            Next i
Next ws
                            
End Sub