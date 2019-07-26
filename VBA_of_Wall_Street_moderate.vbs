Sub stockCalculator()
    'Moderate Solution WITH Challenge
    'Set variables
    Dim i As Double
    Dim currticker As String
    Dim ttlvol As Double
    Dim lastrow As Double
    Dim tickercount As Integer
    Dim ws As Worksheet
    Dim openprice As Double
    Dim closeprice As Double
    
        
    'Loop through each Worksheet
    For Each ws In Worksheets
        'Set the active worksheet
        Dim WorksheetName As String
        WorksheetName = ws.Name
        Sheets(ws.Name).Select
        
        'Set starting point for calculations
        currticker = Range("A2")
        openprice = Range("C2")
        ttlvol = 0
        tickercount = 1
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        'Nxt line used to check lastrow is working correctly, can comment out if good
        'MsgBox ("lastrow =" & lastrow)
    
        'Write out the headers
        Range("I1") = "Ticker"
        Range("J1") = "Yearly Change"
        Range("K1") = "Percent Change"
        Range("L1") = "Total Stock Volume"
    
        For i = 2 To lastrow   'note start at row 2 since 1 is header
            If Cells(i, 1).Value = currticker Then
                'add to volume
                ttlvol = ttlvol + Cells(i, 7).Value
                closeprice = Cells(i, 6).Value
            Else
                'write out the ticker and ttlvolume
                Cells(tickercount + 1, 9) = currticker
                Cells(tickercount + 1, 12) = ttlvol
                Cells(tickercount + 1, 10) = closeprice - openprice
                'If opening price is 0, can't calculate % incr. (div 0 error)
                If openprice = 0 Then
                    Cells(tickercount + 1, 11) = 0
                Else
                    Cells(tickercount + 1, 11) = (closeprice - openprice) / openprice
                End If
                'reset the currticker and ttlvol
                currticker = Cells(i, 1).Value
                ttlvol = Cells(i, 7).Value
                openprice = Cells(i, 3).Value
                'increment the tickercount
                tickercount = tickercount + 1
            End If
        Next i
        'when finished entering totals, do some formatting
        lastrow = Cells(Rows.Count, 9).End(xlUp).Row
        For i = 2 To lastrow
            Cells(i, 11).NumberFormat = "0.00%"
            If Cells(i, 10) > 0 Then
                Cells(i, 10).Interior.Color = vbGreen
            Else
                Cells(i, 10).Interior.Color = vbRed
            End If
        Next i
        Columns("J:L").EntireColumn.AutoFit
    Next ws
    
End Sub

