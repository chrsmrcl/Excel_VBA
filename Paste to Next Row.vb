Sub Button32_Click()
'

'D16 = source data cell
'F = Destination Column
    Dim lastRow As Long
With ActiveSheet
    lastRow = .Cells(.Rows.Count, "F").End(xlUp).Row
End With
'To place data in next blank cell
Cells(lastRow + 1, "F") = Range("D16")
End Sub
