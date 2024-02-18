Sub stock()
    Dim si As Long
    Dim openPrice As Double
    Dim lastRow As Long
    Dim totalVolume As Double
    Dim currentTicker As String
    Dim ticker_open_close_counter As Double
    Dim yearly_open, yearly_close As Double
    Dim ticker As String
    si = 2
    ticker_open_close_counter = 2
    totalVolume = 0
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        totalVolume = totalVolume + Cells(i, 7).Value
        ticker = Cells(i, 1).Value
        yearly_open = Cells(ticker_open_close_counter, 3)
        
        'Summarize if ticker is different
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            yearly_close = Cells(i, 6)
            Cells(si, 9).Value = ticker
            Cells(si, 10).Value = yearly_close - yearly_open
            'set value to null to avoid divinding by zero
            If yearly_open = 0 Then
                Cells(si, 11).Value = Null
            Else
                Cells(si, 11).Value = (yearly_close - yearly_open) / yearly_open
            End If
            Cells(si, 12).Value = totalVolume
            
            ' Color the cells accordinly
            If Cells(si, 10).Value > 0 Then
                Cells(si, 10).Interior.ColorIndex = 4
            Else
                Cells(si, 10).Interior.ColorIndex = 3
            End If
            'Format as percentages
            Cells(si, 11).NumberFormat = "0.00%"
            
            
            totalVolume = 0
            si = si + 1
            ticker_open_close_counter = i + 1
        End If
        
    Next i
    Columns("J").AutoFit
    Columns("K").AutoFit
    Columns("L").AutoFit
End Sub
