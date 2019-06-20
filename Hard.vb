Sub hard()

Dim ws As Worksheet
'Loop through all worksheets
For Each ws In Worksheets
ws.Activate

Dim j As Double
Dim i As Long
Dim k As Long
Dim m As Long
Dim n As Long
Dim x As Long
Dim y As Long
Dim z As Long
Dim Ticker_Symbol As String
Dim TSV As Double
Dim open_price As Double
Dim close_price As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim g_increase As Double
Dim g_decrease As Double
Dim gt_volume As Double
Dim g_increase_ticker As String
Dim g_decrease_ticker As String
Dim gt_ticker As String



'assign column headers
Range("I1").Value = "Ticker Symbol"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"

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

'grab last row value in percent change column
pChange_lastrow = Cells(Rows.Count, 11).End(xlUp).Row

'find greatest percent increase
For k = 2 To pChange_lastrow
    
    If Cells(k, 11).Value > g_increase Then
    g_increase = Cells(k, 11).Value
    g_increase_ticker = Cells(k, 9).Value
    End If
    
    Next k
    
    'assign greatest % increase
    Cells(2, 17).Value = g_increase
    Cells(2, 17).NumberFormat = "0.00%"
    Cells(2, 16).Value = g_increase_ticker
    
'find greatest percent decrease
For m = 2 To pChange_lastrow
    
    If Cells(m, 11).Value < g_decrease Then
    g_decrease = Cells(m, 11).Value
    g_decrease_ticker = Cells(m, 9).Value
    End If
    
    Next m
    
    'assign greatest % decrease
    Cells(3, 17).Value = g_decrease
    Cells(3, 17).NumberFormat = "0.00%"
    Cells(3, 16).Value = g_decrease_ticker
    
'find greatest total volume
For n = 2 To pChange_lastrow
    
    If Cells(n, 12).Value > gt_volume Then
    gt_volume = Cells(n, 12).Value
    gt_ticker = Cells(n, 9).Value
    End If
    
    Next n
    
    'assign greatest % decrease
    Cells(4, 17).Value = gt_volume
    Cells(4, 16).Value = gt_ticker
    
'autofit all columns
Cells.Columns.AutoFit
   
'reset all stored variables for the next sheet
TSV = 0
open_price = 0
close_price = 0
yearly_change = 0
percent_change = 0
g_increase = 0
g_decrease = 0
gt_volume = 0
Ticker_Symbol = Empty
g_increase_ticker = Empty
g_decrease_ticker = Empty
gt_ticker = Empty

'next worksheet
Next

End Sub
