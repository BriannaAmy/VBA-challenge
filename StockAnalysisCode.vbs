Attribute VB_Name = "Module1"
Sub StockAnalysis():
    
    For Each ws In Worksheets
    
        Dim rowCount As Long
        Dim totVol As Double
        Dim ticker As String
        Dim yearlyChg As Double
        Dim percentChg As Double
        Dim openPrice As Double
        Dim summaryRow As Long
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        summaryRow = 2
        totVol = 0
        yearlyChg = 0
        rowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row 'getting last row number
        
        For Row = 2 To rowCount
        
            If ws.Cells(Row, 3).Value <> 0 Then 'check for first open price for first ticker
            
                openPrice = ws.Cells(Row, 3).Value
                Exit For 'exit loop after first open price found
                
            End If
            
        Next Row
        
        For Row = 2 To rowCount
        
            If ws.Cells(Row, 1).Value <> ws.Cells(Row + 1, 1).Value Then 'checking for change in ticker name
                
                'pull data for summary table
                ticker = ws.Cells(Row, 1).Value 'pull ticker name
                totVol = totVol + ws.Cells(Row, 7).Value 'add onto total volume one last time
                yearlyChg = ws.Cells(Row, 6).Value - openPrice
                percentChg = yearlyChg / openPrice
                
                'summary table population
                ws.Cells(summaryRow, 9).Value = ticker
                ws.Cells(summaryRow, 10).Value = yearlyChg
                ws.Cells(summaryRow, 11).Value = percentChg
                ws.Cells(summaryRow, 12).Value = totVol
                
                ws.Cells(summaryRow, 11).NumberFormat = "0.00%" 'format percentage change
                
                'reset variables
                openPrice = ws.Cells(Row + 1, 3).Value
                totVol = 0
                summaryRow = summaryRow + 1
            
            ElseIf openPrice <> 0 Then 'check the opening price to make it does not equal 0
                
                totVol = totVol + ws.Cells(Row, 7).Value 'running total for total volume of a ticker
                
            Else
            
                openPrice = ws.Cells(Row, 3).Value
                
            End If
        
        Next Row
        
        rowCount = ws.Cells(Rows.Count, "I").End(xlUp).Row 'getting last row number for summary table
        
        For Row = 2 To rowCount
        
            If ws.Cells(Row, 10) < 0 Then
            
                ws.Cells(Row, 10).Interior.ColorIndex = 3
                
            ElseIf ws.Cells(Row, 10) > 0 Then
            
                ws.Cells(Row, 10).Interior.ColorIndex = 4
            
            End If
        
        Next Row
        
        Dim maxChg As Double
        Dim maxChgTicker As String
        Dim minChg As Double
        Dim minChgTicker As String
        Dim maxVol As Double
        Dim maxVolTicker As String
        
        'max percent change
        maxChg = WorksheetFunction.max(ws.Range("K2:K" & rowCount))
        maxChgTicker = WorksheetFunction.Index(ws.Range("I2:L" & rowCount), WorksheetFunction.Match(maxChg, ws.Range("K2:K" & rowCount), 0), 1)
        'min percent change
        minChg = WorksheetFunction.min(ws.Range("K2:K" & rowCount))
        minChgTicker = WorksheetFunction.Index(ws.Range("I2:L" & rowCount), WorksheetFunction.Match(minChg, ws.Range("K2:K" & rowCount), 0), 1)
        'max total volume
        maxVol = WorksheetFunction.max(ws.Range("L2:L" & rowCount))
        maxVolTicker = WorksheetFunction.Index(ws.Range("I2:L" & rowCount), WorksheetFunction.Match(maxVol, ws.Range("L2:L" & rowCount), 0), 1)
        
        'populate min/max table
        ws.Cells(2, 16).Value = maxChgTicker
        ws.Cells(2, 17).Value = maxChg
        
        ws.Cells(2, 17).NumberFormat = "0.00%"
        
        ws.Cells(3, 16).Value = minChgTicker
        ws.Cells(3, 17).Value = minChg
        
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
        ws.Cells(4, 16).Value = maxVolTicker
        ws.Cells(4, 17).Value = maxVol
    
    Next ws

End Sub
