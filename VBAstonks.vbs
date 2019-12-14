Dim ticker As String
Dim stockVolume As Double
Dim TableRow As Double
Dim yearlyChange As Double
Dim percentChange As Double
Dim initialStock As Double
Dim closingStock As Double
TableRow = 2

'titles
Cells(1, 9) = "Ticker"
Cells(1, 10) = "Yearly Change"
Cells(1, 11) = "Percent Change"
Cells(1, 12) = "Total Stock Volume"

'loops throught worksheets
For Each ws In Worksheets

    'finds last row
    lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastRow
    
        'finds first stock value for each individual ticker string when switching ticker
        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
            initialStock = ws.Cells(i, 3)
            
        'finds, posts, and resets variables to be posted in summary area.
        ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            closingStock = ws.Cells(i, 6)
            'adds stock value on final row of ticker
            stockVolume = stockVolume + ws.Cells(i, 7).Value
            yearlyChange = closingStock - initialStock
            
            'filters out 0 values to prevent overflow errors
            If initialStock = 0 Or closingStock = 0 Then
                percentageChange = Null
            Else
                'calculates percentage change
                percentChange = yearlyChange / initialStock
            End If
            
            ticker = ws.Cells(i, 1).Value
            'posts acquired values
            Range("I" & TableRow).Value = ticker
            Range("J" & TableRow).Value = yearlyChange
                                        'formats percentChange decimal value as a percentage
            Range("K" & TableRow).Value = Format(percentChange, "0.00%")
            Range("L" & TableRow).Value = stockVolume
            'moves summary table down one unit
            TableRow = TableRow + 1
            'resets values
            stockVolume = 0
            yearlyChange = 0
            
        'adds to stock volume when ticker below matches ticker above
        ElseIf ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
            stockVolume = stockVolume + ws.Cells(i, 7).Value
        End If
        
        'sets conditional formatting
        If Cells(i, 10).Value > 0 Then
            Cells(i, 10).Interior.ColorIndex = 4
        ElseIf Cells(i, 10).Value < 0 Then
            Cells(i, 10).Interior.ColorIndex = 3
        End If
    Next i
Next ws
End Sub
