Dim TableRow As Double
Dim stockChange As Double
Dim yearlyChange As Double
Dim percentChange As Double
Dim initialStock As Double
Dim closingStock As Double
TableRow = 2

Cells(1, 9) = "Ticker"
Cells(1, 10) = "Yearly Change"
Cells(1, 11) = "Percent Change"
Cells(1, 12) = "Total Stock Volume"

For Each ws In Worksheets
    lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastRow
        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
            initialStart = ws.Cells(i, 3)
        ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            closingStock = ws.Cells(i, 6)
            stockVolume = stockVolume + ws.Cells(i, 7).Value
            stockChange = ws.Cells(i, 6).Value - ws.Cells(i, 3).Value
            yearlyChange = yearlyChange + stockChange
            ticker = ws.Cells(i, 1).Value
            Range("I" & TableRow).Value = ticker
            Range("J" & TableRow).Value = yearlyChange
            Range("L" & TableRow).Value = stockVolume
            TableRow = TableRow + 1
            stockVolume = 0
            yearlyChange = 0
        ElseIf ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
            stockVolume = stockVolume + ws.Cells(i, 7).Value
            stockChange = ws.Cells(i, 6).Value - ws.Cells(i, 3).Value
            yearlyChange = yearlyChange + stockChange
        End If
    Next i
Next ws
End Sub
