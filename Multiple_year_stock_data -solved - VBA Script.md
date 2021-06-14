
Sub stock_analysis()
'First I set all dimensions, values and ws set up to apply all codes in all worksheets.

'Dimensions

Dim total As Double
Dim change As Double
Dim percentChange As Double
Dim dailyChange As Double
Dim averageChange As Double
Dim increasenumber As Double
Dim decreasenumber As Double
Dim volumenumber As Long
Dim start As Long
Dim rowcount As Long
Dim i As Long
Dim j As Integer
Dim days As Integer
Dim ws As Worksheet
For Each ws In Worksheets


'values (set up values from 0 as there is no value, only start 2 since there is a values start in row 2.

start = 2
j = 0
total = 0
change = 0

'New row name

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "yearly_change"
ws.Range("K1").Value = "percent_change"
ws.Range("L1").Value = "total_stock_volume"

'New Row and column name for result part
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest %Increase"
ws.Range("O3").Value = "Greatest %Decrease"
ws.Range("O4").Value = "Greatest Total Volume"


'rowcount for spreadsheet which is required for all calculation.

rowcount = ws.Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To rowcount

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        total = total + ws.Cells(i, 7).Value

                If total = 0 Then
                ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                ws.Range("J" & 2 + j).Value = 0
                ws.Range("K" & 2 + j).Value = "%" & 0
                ws.Range("L" & 2 + j).Value = 0

                 Else
        
                 If ws.Cells(start, 3) = 0 Then
                 For find_value = start To i
                 
                 If ws.Cells(find_value, 3).Value <> 0 Then
                 start = find_value

                 Exit For

                 End If

                 Next find_value
            End If

'find percentchange
    change = (ws.Cells(i, 6) - ws.Cells(start, 3))
    percentChange = Round((change / ws.Cells(start, 3) * 100), 2)


'find start of the next ticker

    start = i + 1
    ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
    ws.Range("J" & 2 + j).Value = Round(change, 2)
    ws.Range("K" & 2 + j).Value = "%" & percentChange
    ws.Range("L" & 2 + j).Value = total

'formatting (positives green and negatives red)
    If ws.Range("J" & 2 + j) > 0 Then
    ws.Range("J" & 2 + j).Interior.ColorIndex = 4
 
    ElseIf ws.Range("J" & 2 + j) < 0 Then
    ws.Range("J" & 2 + j).Interior.ColorIndex = 3

    Else
    Range("J" & 2 + j).Interior.ColorIndex = 0

    End If


End If


'RESET
        total = 0
        change = 0
        j = j + 1
        
        Else
        total = total + ws.Cells(i, 7).Value

End If

Next i
 
ws.Range("P1").Value = "Ticker"
    
        ws.Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & rowcount)) * 100
        ws.Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & rowcount)) * 100
        ws.Range("Q4") = WorksheetFunction.Max(Range("L2:L" & rowcount))
        
        increasenumber = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowcount)), ws.Range("K2:K" & rowcount), 0)
        decreasenumber = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowcount)), ws.Range("K2:K" & rowcount), 0)
        volumenumber = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowcount)), ws.Range("L2:L" & rowcount), 0)
        
        ws.Range("P2") = ws.Cells(increasenumber + 1, 9)
        ws.Range("P3") = ws.Cells(decreasenumber + 1, 9)
        ws.Range("P4") = ws.Cells(volumenumber + 1, 9)
        

Next ws

End Sub