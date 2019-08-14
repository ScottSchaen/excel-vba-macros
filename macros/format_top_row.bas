Sub FormatTopRow()
'Purpose: Freezes and formats the top row of your table to make it easier to look at and work with
'Active sheet only
Dim toprow As Range
'If there's <=1 used cell in row 1 then check if the active cell is inside the table to format
If Application.WorksheetFunction.CountA(Range("1:1")) > 1 Then
    Set toprow = Range("A1:" & Range("IV1").End(xlToLeft).Address)
Else
    Dim tbl As Range
    Set tbl = Selection.CurrentRegion
    'If the active cell is not inside of a table then inform user and end macro
    If tbl.Count = 1 Then
        MsgBox "Couldn't find a table to format! Click a cell in the table and run again", vbExclamation, "Couldn't find table!"
        Exit Sub
    End If
    Dim firstcell As Range
    Set firstcell = tbl.Cells(1, 1)
    Set toprow = Range(firstcell, firstcell.Offset(0, tbl.Columns.Count - 1))
End If
Cells(toprow.Row + 1, 1).Select
ActiveWindow.FreezePanes = False
ActiveWindow.FreezePanes = True
'Sets a grey background with white bold text
With toprow.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .ThemeColor = xlThemeColorDark2
    .TintAndShade = -0.249977111117893
    .PatternTintAndShade = 0
End With
toprow.Font.Bold = True
toprow.Font.Color = vbWhite
End Sub
