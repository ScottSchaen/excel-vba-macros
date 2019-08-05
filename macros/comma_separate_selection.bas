Sub CommaSeparateSelection()
'Purpose: Comma separates all cells in selection and outputs them to an unused adjacent cell
'Current sheet only
Dim outputcell As Range
Set outputcell = Range("IV1").End(xlToLeft).Offset(0, 1)
'Wrap comma separated values in quotes yes/no
apos = MsgBox("Add apostrophes?", vbYesNo, "Add apostrophes and wrap selections in quotes?")
If apos = vbYes Then apos = True Else apos = False
For Each cell In Selection
    If cell.Value <> "" Then
        If apos = False Then outputcell.Value = outputcell.Value & cell.Value & ", "
        If apos = True Then outputcell.Value = outputcell.Value & "'" & cell.Value & "', "
    End If
Next cell
'Removes trailing comma
outputcell.Value = Left(outputcell.Value, Len(outputcell.Value) - 2)
End Sub
