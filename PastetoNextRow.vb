Sub Button32_Click()
'
' Button32_Click Macro
'

'
    Range("D17").Select
    Selection.Copy
    Range("F21").Select
'
    Dim lastRow As Long
With ActiveSheet
    lastRow = .Cells(.Rows.Count, "F").End(xlUp).Row
End With
'To place data in next blank cell
Cells(lastRow + 1, "F") = Range("D17")
End Sub
