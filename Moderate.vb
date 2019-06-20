Sub Moderate()

Dim ws As Worksheet
'Loop through all worksheets
For Each ws In Worksheets
ws.Activate

Dim i As Long
Dim x As Long
Dim y As Long
Dim z As Long
Dim Ticker_Symbol As String
Dim TSV As Double
Dim open_price As Double
Dim close_price As Double
Dim yearly_change As Double
Dim percent_change As Double

'assign column headers
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

'grab last row value
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'assign initial variables
x = 2
y = 2
z = 2
TSV = 0
open_price = Range("C2").Value

For i = 2 To lastrow
    
    If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
    TSV = TSV + Cells(i, 7).Value
    open_price = open_price
    Else
    'final addition to the total stock price
    TSV = TSV + Cells(i, 7).Value
    
    Ticker_Symbol = Cells(i, 1).Value
    Cells(x, 9).Value = Ticker_Symbol
    Cells(y, 12).Value = TSV
    
    'assign closing price for the last day of the year
    close_price = Cells(i, 6).Value
    
    'calculate/assign yearly change
    yearly_change = close_price - open_price
    Cells(z, 10).Value = yearly_change
    Cells(z, 10).NumberFormat = "0.000000000"
        'assign cell color based on relation to zero
        If Cells(z, 10).Value > 0 Then
        Cells(z, 10).Interior.Color = RGB(0, 255, 0)
        ElseIf Cells(z, 10).Value < 0 Then
        Cells(z, 10).Interior.Color = RGB(255, 0, 0)
        ElseIf Cells(z, 10).Value = 0 Then
        Cells(z, 10).Interior.Color = RGB(255, 255, 255)
        End If
    
    
    'calculate percent change
    'Protect against division by zero
    If open_price = 0 Then
    Cells(z, 11).Value = "0"
    Cells(z, 11).NumberFormat = "0.00%"
    Else
    percent_change = (yearly_change / open_price)
    Cells(z, 11).Value = percent_change
    Cells(z, 11).NumberFormat = "0.00%"
    End If
    
    'reset open price for the next ticker symbol
    open_price = Cells(i + 1, 3).Value
    
    'variables adjusted
    x = x + 1
    y = y + 1
    z = z + 1
    TSV = 0
    
    End If
    
Next i

'autofit all columns
Cells.Columns.AutoFit
   
'reset all stored variables for the next sheet
TSV = 0
open_price = 0
close_price = 0
yearly_change = 0
percent_change = 0
Ticker_Symbol = Empty

'next worksheet
Next

End Sub
