Sub mulitiple_year()

'Define variables
Dim i As Long, j As Integer
Dim Column As Integer
Column = 1
Dim rowCount As Long
Dim Total As Double
Dim Start As Long
Dim Change As Double
Dim percentChange As Double
Dim days As Integer
Dim dailyChange As Double
Dim averageChange As Double
Dim ws As Worksheet

'Create loop for every worksheet
For Each ws in Worksheets

'Label the columns with programming
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

j = 0
Total = 0
Change = 0
Start = 2

 'get the row number of the last row with data
 rowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row

'Loop through the entire worksheet
For i = 2 To rowCount

   If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
     Total = Total + ws.Cells(i, 7).Value
     If Total = 0 Then
     ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
     ws.Range("J" & 2 + j).Value = 0
     ws.Range("K" & 2 + j).Value = "%" & 0
     ws.Range("L" & 2 + j).Value = 0

     Else
     If ws.Cells(Start, 3) = 0 Then
     For find_value = Start To i
     If ws.Cells(find_value, 3).Value <> 0 Then
     Start = find_value
     Exit For
     End If
     Next find_value
     End If

     Change = (ws.Cells(i, 6) - ws.Cells(Start, 3))
     percentChange = Round((Change / ws.Cells(Start, 3) * 100), 2)

     Start = i + 1

     ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
     ws.Range("J" & 2 + j).Value = Round(Change, 2)
     ws.Range("K" & 2 + j).Value = "%" & percentChange
     ws.Range("L" & 2 + j).Value = Total
 
 'Highlight the cells for green = positive, red = negative
 Select Case Change
 Case Is > 0
 ws.Range("j" & 2 + j).Interior.ColorIndex = 4
 Case Is < 0
 ws.Range("j" & 2 + j).Interior.ColorIndex = 3
 Case Else
 ws.Range("j" & 2 + j).Interior.ColorIndex = 0

 End Select

 End If

 Total = 0
 Change = 0
 j = j + 1
 days = 0

 Else

 Total = Total + ws.Cells(i, 7).Value

 End If
 Next i

'Bonus - Greatest % increase, decrease and total volume
ws.Cells(2, 17).Value = "%" & WorksheetFunction.Max(ws.Range("K2:K" & rowCount)) * 100
ws.Cells(3, 17).Value = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowCount)) * 100
ws.Cells(4, 17).Value = WorksheetFunction.Max(ws.Range("L2:L" & rowCount))

increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:k" & rowCount)), ws.Range("K2:K" & rowCount), 0)
decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:k" & rowCount)), ws.Range("K2:K" & rowCount), 0)
volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCount)), ws.Range("L2:L" & rowCount), 0)

ws.Cells(2, 16).Value = ws.Cells(increase_number + 1, 9)
ws.Cells(3, 16).Value = ws.Cells(decrease_number + 1, 9)
ws.Cells(4, 16).Value = ws.Cells(volume_number + 1, 9)

Next ws

End Sub

     


