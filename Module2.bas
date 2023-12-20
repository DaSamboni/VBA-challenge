Attribute VB_Name = "Module1"
Sub StocksProject()


'Set up multiple worksheets
Dim ws As Worksheet

For Each ws In Worksheets


'Fill column-toppers with necessary text
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"


'Define used variables
Dim Ticker_Name As String
Ticker_Name = " "

Dim Stock_Volume As Double
Stock_Volume = 0

Dim Daily_Change As Double
Daily_Change = 0

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

Dim Ticker_Row As Integer
Ticker_Row = 1

Dim Open_Value As Double
Open_Value = ws.Cells(2, 3).Value

Dim Lastrow As Long
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Dim MaxP As Double
MaxP = 0

Dim MinP As Double
MinP = 0

Dim MaxV As Double
MaxV = 0



'Begin For loop
For i = 2 To Lastrow


'Stop loop after Ticker value changes
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

Ticker_Row = Ticker_Row + 1

Ticker_Name = ws.Cells(i, 1).Value

ws.Cells(Ticker_Row, "I").Value = Ticker_Name

Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value

Daily_Change = Daily_Change + (ws.Cells(i, 6).Value - ws.Cells(i, 3).Value)

ws.Cells(Summary_Table_Row, 9).Value = Ticker_Name

ws.Cells(Summary_Table_Row, 12).Value = Stock_Volume

ws.Cells(Summary_Table_Row, 10).Value = Daily_Change

    'Nested If
    If ws.Cells(Summary_Table_Row, 10).Value > 0 Then
    
    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
    
    ElseIf ws.Cells(Summary_Table_Row, 10).Value < 0 Then
    
    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
    
    End If
    
ws.Cells(Summary_Table_Row, 11).Value = (Daily_Change / Open_Value)

Summary_Table_Row = Summary_Table_Row + 1

Stock_Volume = 0

Daily_Change = 0

Open_Value = ws.Cells(i + 1, 3).Value

Else

Daily_Change = Daily_Change + (ws.Cells(i, 6).Value - ws.Cells(i, 3).Value)

Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value

End If


Next i

ws.Range("K2", "K10000").NumberFormat = "0.00%"

MaxP = Application.WorksheetFunction.Max(ws.Range("K:K"))

MinP = Application.WorksheetFunction.Min(ws.Range("K:K"))

MaxV = Application.WorksheetFunction.Max(ws.Range("L:L"))

ws.Cells(2, 17) = MaxP

For j = 2 To Lastrow

ws.Cells(3, 17) = MinP

ws.Cells(4, 17) = MaxV


'Filling Max,Min, and Volume with Ticker Values
If ws.Cells(j, 11).Value = MaxP Then

    ws.Cells(2, 16).Value = ws.Cells(j, 9).Value

    End If

If ws.Cells(j, 11).Value = MinP Then

    ws.Cells(3, 16).Value = ws.Cells(j, 9).Value

    End If

If ws.Cells(j, 12).Value = MaxV Then

    ws.Cells(4, 16).Value = ws.Cells(j, 9).Value

    End If

Next j

ws.Range("Q2", "Q3").NumberFormat = "0.00%"

Next ws


End Sub


