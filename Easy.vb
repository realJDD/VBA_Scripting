Sub Easy()

Dim ws As Worksheet
'Loop through all worksheets
For Each ws In Worksheets
ws.Activate

Dim i As Long
Dim x As Long
Dim y As Long
Dim Ticker_Symbol As String
Dim TSV As Double

'assign column headers
Range("I1").Value = "Ticker Symbol"
Range("J1").Value = "Total Stock Volume"

'grab last row value
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'assign initial variables
x = 2
y = 2
TSV = 0

For i = 2 To lastrow

    If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
    TSV = TSV + Cells(i, 7).Value
    Else
    'final addition to the total stock price
    TSV = TSV + Cells(i, 7).Value
    Ticker_Symbol = Cells(i, 1).Value
    Cells(x, 9).Value = Ticker_Symbol
    Cells(y, 10).Value = TSV

    'variables adjusted
    x = x + 1
    y = y + 1
    TSV = 0

    End If

Next i

'autofit all columns
Cells.Columns.AutoFit

'reset all stored variables for the next sheet
TSV = 0
Ticker_Symbol = Empty

'next worksheet
Next

End Sub
