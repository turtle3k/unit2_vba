Sub StockCalculator_Easy()
    'Easy Solution WITH Challenge
    'Set variables
    Dim i As Double
    Dim currticker As String
    Dim ttlvol As Double
    Dim lastrow As Double
    Dim tickercount As Integer
    Dim ws As Worksheet
    
        
    'Loop through each Worksheet
    For Each ws In Worksheets
        'Set the active worksheet
        Dim WorksheetName As String
        WorksheetName = ws.Name
        Sheets(ws.Name).Select
        
        'Set starting point for calculations
        currticker = Range("A2")
        ttlvol = 0
        tickercount = 1
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        'Nxt line used to check lastrow is working correctly, can comment out if good
        'MsgBox ("lastrow =" & lastrow)
    
        'Write out the headers
        Range("I1") = "Ticker"
        Range("J1") = "Total Stock Volume"
    
        For i = 2 To lastrow   'note start at row 2 since 1 is header
            If Cells(i, 1).Value = currticker Then
                'add to volume
                ttlvol = ttlvol + Cells(i, 7).Value
            Else
                'write out the ticker and ttlvolume
                Cells(tickercount + 1, 9) = currticker
                Cells(tickercount + 1, 10) = ttlvol
                'reset the currticker and ttlvol
                currticker = Cells(i, 1).Value
                ttlvol = Cells(i, 7).Value
                'increment the tickercount
                tickercount = tickercount + 1
            End If
        Next i
        'when finished entering the volume totals, autoset the new column width
        Columns("J:J").EntireColumn.AutoFit
    Next ws
    
End Sub
