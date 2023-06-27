Attribute VB_Name = "Module1"
Sub Analysis()


Cells(1, 9) = "Ticker"
Range("A2").Select
Tracker = ActiveCell
Counter = StrComp(ActiveCell, Tracker)
spot = 2

Do Until IsEmpty(ActiveCell)

If Counter = 0 Then
ActiveCell.Offset(1, 0).Select
Counter = StrComp(ActiveCell, Tracker)
Else
Cells(spot, 9) = Tracker
Tracker = ActiveCell
ActiveCell.Offset(1, 0).Select
Counter = StrComp(ActiveCell, Tracker)
spot = spot + 1
End If



Loop



Cells(1, 10) = "Yearly Change"
Range("A2").Select
Tracker = ActiveCell
Counter = StrComp(ActiveCell, Tracker)
Counter2 = 0
spot = 2

Do Until IsEmpty(ActiveCell)

If Counter = 0 Then
Counter2 = Counter2 + (Cells(ActiveCell.Row, 6) - Cells(ActiveCell.Row, 3))
ActiveCell.Offset(1, 0).Select
Counter = StrComp(ActiveCell, Tracker)
Else
Cells(spot, 10) = Counter2
If Counter2 < 0 Then
Cells(spot, 10).Interior.ColorIndex = 3
Else
Cells(spot, 10).Interior.ColorIndex = 4
End If
Counter2 = 0
Tracker = ActiveCell
ActiveCell.Offset(1, 0).Select
Counter = StrComp(ActiveCell, Tracker)
spot = spot + 1
End If



Loop


Cells(1, 11) = "Percent Change"
Range("A2").Select
Tracker = ActiveCell
Counter = StrComp(ActiveCell, Tracker)
Counter2 = 0
OldValue = 0
PChange = 0
spot = 2

Do Until IsEmpty(ActiveCell)

If Counter = 0 Then
Counter2 = Counter2 + (Cells(ActiveCell.Row, 6) - Cells(ActiveCell.Row, 3))
OldValue = OldValue + Cells(ActiveCell.Row, 3)
ActiveCell.Offset(1, 0).Select
Counter = StrComp(ActiveCell, Tracker)
Else
PChange = (Counter2 / OldValue) * 100
Cells(spot, 11) = PChange
Cells(spot, 11).NumberFormat = "0.00%"
Counter2 = 0
OldValue = 0
PChange = 0
Tracker = ActiveCell
ActiveCell.Offset(1, 0).Select
Counter = StrComp(ActiveCell, Tracker)
spot = spot + 1
End If



Loop



Cells(1, 12) = "Total Stock volume"
Range("A2").Select
Tracker = ActiveCell
Counter = StrComp(ActiveCell, Tracker)
Counter2 = 0
spot = 2

Do Until IsEmpty(ActiveCell)

If Counter = 0 Then
Counter2 = Counter2 + Cells(ActiveCell.Row, 7)
ActiveCell.Offset(1, 0).Select
Counter = StrComp(ActiveCell, Tracker)
Else
Cells(spot, 12) = Counter2
Counter2 = 0
Tracker = ActiveCell
ActiveCell.Offset(1, 0).Select
Counter = StrComp(ActiveCell, Tracker)
spot = spot + 1
End If



Loop






End Sub



