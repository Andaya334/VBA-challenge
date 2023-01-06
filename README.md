Sub Stockdata()

Dim ws As Worksheet
Dim y As Double
Dim x As Double
Dim lastrow As Double
Dim yropen As Single
Dim yrclose As Single
Dim volume As Double

For Each ws In Sheets
Worksheets(ws.Name).Activate
y = 2
x = 2
lastrow = WorksheetFunction.CountA(ActiveSheet.Columns(1))
volume = 0

'Use Loop to find unique tickers, and place in respective columns
For i = 2 To lastrow
ticker = Cells(i, 1).Value
previous_ticker = Cells(i - 1, 1).Value
If ticker <> previous_ticker Then
Cells(x, 9).Value = ticker
x = x + 1
End If
Next i

'Set the column names in their respective cells
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("O2").Value = "Greatest % increase"
Range("O3").Value = "Greatest % decrease"
Range("O4").Value = "Greatest total volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"


'Use loop for total stock volume and yearly change and percentage
For i = 2 To lastrow
ticker = Cells(i, 1).Value
previous_ticker = Cells(i - 1, 1).Value
next_ticker = Cells(i + 1, 1).Value
If ticker <> previous_ticker Then
    yropen = Cells(i, 3).Value
    volume = Cells(i, 7).Value
    
ElseIf ticker <> next_ticker Then
    yrclose = Cells(i, 6).Value
    increase = yrclose - yropen
    Percent = increase / yropen
    volume = volume + Cells(i, 7).Value
    Cells(y, 10).Value = increase
    Cells(y, 11).Value = FormatPercent(Percent)
    Cells(y, 12).Value = volume
    yropen = 0
    yrclose = 0
    y = y + 1
    
Else
    volume = volume + Cells(i, 7).Value
End If
Next i

'Now color code either green or red yearly change
For i = 2 To 3001
If Cells(i, 10).Value > 0 Then
Cells(i, 10).Interior.ColorIndex = 4
Else
Cells(i, 10).Interior.ColorIndex = 3
End If
Next i

'Max, min, and volume of values and put into correct cells
Max = WorksheetFunction.Max(Range("K2:K3001"))
Min = WorksheetFunction.Min(Range("K2:K3001"))
vol = WorksheetFunction.Max(Range("L2:L3001"))

Range("Q2").Value = FormatPercent(Max)
Range("Q3").Value = FormatPercent(Min)
Range("Q4").Value = FormatPercent(vol)

For i = 2 To lastrow
If Max = Cells(i, 11).Value Then
Range("P2").Value = Cells(i, 9).Value
ElseIf Min = Cells(i, 11).Value Then
Range("P3").Value = Cells(i, 9).Value
ElseIf vol = Cells(i, 12).Value Then
Range("P4").Value = Cells(i, 9).Value
End If

Next i

Next ws

End Sub

