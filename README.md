# Stock-Data-Analysis
Week 2 Module Challenge 
'VBA CODE

Sub Stock_Data()

Dim Total As Double
Dim Row As Long
Dim Change As Double
Dim Column As Integer
Dim Start As Long
Dim Count As Long
Dim PercentChange As Double
Dim Days As Integer
Dim DailyChange As Single
Dim AverageChange As Double
Dim WS As Worksheet

For Each WS In Worksheets
    Column = 0
    Total = 0
    Start = 2
    Change = 0

WS.Range("i1").Value = "Ticker"
WS.Range("J1").Value = "Yearly Change"
WS.Range("K1").Value = "Percent Change"
WS.Range("L1").Value = "Total Stock Volume"
WS.Range("P1").Value = "Ticker"
WS.Range("Q1").Value = "Value"
WS.Range("O2").Value = "Greatest % Increase"
WS.Range("O3").Value = "Greatest % Decrease"
WS.Range("O4").Value = "Greatest Total Volume"

Count = WS.Cells(Rows.Count, "A").End(xlUp).Row

For Row = 2 To Count
    If WS.Cells(Row + 1, 1).Value <> WS.Cells(Row, 1).Value Then
    
    Total = Total + WS.Cells(Row, 7).Value
    
    If Total = 0 Then
    
    WS.Range("i" & 2 + Column).Value = Cells(Row, 1).Value
    WS.Range("J" & 2 + Column).Value = 0
    WS.Range("K" & 2 + Column).Value = "%" & 0
    WS.Range("L" & 2 + Column).Value = 0
Else
    If WS.Cells(Start, 3) = 0 Then
    For Find_Value = Start To Row
    If WS.Cells(Find_Value, 3).Value <> 0 Then
    Start = Find_Value
    Exit For
End If
    Next Find_Value
    End If
    
    Change = (WS.Cells(Row, 6) - WS.Cells(Start, 3))
    PercentChange = Change / WS.Cells(Start, 3)
    Start = Row + 1
    WS.Range("i" & 2 + Column) = WS.Cells(Row, 1).Value
    WS.Range("J" & 2 + Column) = Change
    WS.Range("J" & 2 + Column).NumberFormat = "0.00"
    WS.Range("K" & 2 + Column).Value = PercentChange
    WS.Range("K" & 2 + Column).NumberFormat = "0.00%"
    WS.Range("L" & 2 + Column).Value = Total
    
Select Case Change
    Case Is > 0
    WS.Range("J" & 2 + Column).Interior.ColorIndex = 4
    Case Is < 0
    WS.Range("J" & 2 + Column).Interior.ColorIndex = 3
    Case Else
    WS.Range("J" & 2 + Column).Interior.ColorIndex = 0
End Select


End If
    Total = 0
    Change = 0
    Column = Column + 1
    Days = 0
    DailyChange = 0
    
Else
    Total = Total + WS.Cells(Row, 7).Value
    
End If


Next Row
 
    WS.Range("Q2") = "%" & WorksheetFunction.Max(WS.Range("K2:K" & Count)) * 100
    WS.Range("Q3") = "%" & WorksheetFunction.Min(WS.Range("K2:K" & Count)) * 100
    WS.Range("Q4") = WorksheetFunction.Max(WS.Range("L2:L" & Count))
    
    Incrase_number = WorksheetFunction.Match(WorksheetFunction.Max(WS.Range("K2:K" & Count)), WS.Range("K2:K" & Count), 0)
    Decrease_Number = WorksheetFunction.Match(WorksheetFunction.Min(WS.Range("K2:K" & Count)), WS.Range("K2:K" & Count), 0)
    Volume_Number = WorksheetFunction.Match(WorksheetFunction.Max(WS.Range("L2:L" & Count)), WS.Range("L2:L" & Count), 0)
    
    WS.Range("P2") = WS.Cells(Increase_Number + 1, 9)
    WS.Range("P3") = WS.Cells(Decrease_Number + 1, 9)
    WS.Range("P4") = WS.Cells(Volume_Number + 1, 9)
    



Next WS




End Sub
