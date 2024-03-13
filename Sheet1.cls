VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock_market_analysis()
For Each ws In Worksheets


Dim Workshhetnam As String
Dim Ticker As String
Dim Ticker_Name As String
Dim Percent_Change As Double
Dim Total_Stock_Volume As Double
Dim Yearly_Change As Double
Dim Summary_Table_Row As Double
Dim greatest_total As Double
Dim greatest_decrease As Double
Dim greatest_increase As Double
Dim j As Long
Dim i As Long

worksheetname = ws.Name

ws.Cells(1, "J").Value = "Ticker"
ws.Cells(1, "K").Value = "Yearly Change"
ws.Cells(1, "L").Value = "Percent Change"
ws.Cells(1, "M").Value = "Total Stock Volume"
  

Summary_Table_Row = 2
j = 2
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To LastRow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    Ticker_Name = ws.Cells(i, 1).Value
    ws.Range("J" & Summary_Table_Row).Value = Ticker_Name
    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, "G").Value
    ws.Range("M" & Summary_Table_Row).Value = Total_Stock_Volume
    Summary_Table_Row = Summary_Table_Row + 1
    Total_Stock_Volume = 0
    Else
    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, "G").Value
    End If
    Next i



Summary_Table_Row = 2
j = 2
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To LastRow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    Yearly_Change = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
    ws.Range("K" & Summary_Table_Row).Value = Yearly_Change
    Percent_Change = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
    ws.Range("L" & Summary_Table_Row).Value = Percent_Change
    ws.Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
    Summary_Table_Row = Summary_Table_Row + 1
    j = i + 1
    End If
    Next i

Summary_Table_Row = 2
LastRow = ws.Cells(Rows.Count, "J").End(xlUp).Row
    
    For i = 2 To LastRow
    If ws.Cells(i, "K").Value > 0 Then
    ws.Cells(i, "K").Interior.ColorIndex = 4
    Else
    ws.Cells(i, "K").Interior.ColorIndex = 3
    End If
    Next i

ws.Cells(2, "P").Value = "Greatest % increase"
ws.Cells(3, "P").Value = "Greatest % decrease"
ws.Cells(4, "P").Value = "Greatest total volume"
ws.Cells(1, "Q").Value = "ticker"
ws.Cells(1, "R").Value = "value"
LastRowJ = Cells(Rows.Count, "J").End(xlUp).Row

    For i = 2 To LastRow
    If ws.Cells(i, "L").Value > greatest_increase Then
    greatest_increase = ws.Cells(i, "L").Value
    ws.Cells(2, "Q").Value = ws.Cells(i, "J").Value
    Else
    greatest_increase = greatest_increase
    End If
    If ws.Cells(i, "L").Value < greatest_decrease Then
    greatest_decrease = ws.Cells(i, "L").Value
    ws.Cells(3, "Q").Value = ws.Cells(i, "J").Value
    Else
    greatest_decrease = greatest_decrease
    End If
    If ws.Cells(i, "M").Value > greatest_total Then
    greatest_total = ws.Cells(i, "M").Value
    ws.Cells(4, "Q").Value = ws.Cells(i, "J").Value
    Else
    greatest_total = greatest_total
    End If
    ws.Cells(2, "R").Value = greatest_increase
    ws.Cells(2, "R").NumberFormat = "0.00%"
    ws.Cells(3, "R").Value = greatest_decrease
    ws.Cells(3, "R").NumberFormat = "0.00%"
    ws.Cells(4, "R").Value = greatest_total
    ws.Cells(4, "R").NumberFormat = "0"
    Next i
    
Next ws

End Sub








