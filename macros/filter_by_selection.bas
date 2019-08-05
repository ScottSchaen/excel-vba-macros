Sub FilterBySelection()
'Purpose: Filter for current cell in selection
'Only filters one cell so this reduces the selection to one cell if multiple are selected
If Selection.Count > 1 Then ActiveCell.Select
On Error GoTo err
    'Try filtering to selected
    Selection.AutoFilter Field:=Selection.Column, Criteria1:="=" & Selection.Value
Exit Sub
err:
    If err = 1004 Then
        'Turn on autofilter if it's not on already
        Selection.AutoFilter
        Selection.AutoFilter Field:=Selection.Column, Criteria1:="=" & Selection.Value
    Else
        'If it doesn't work, filter to '#N/A'
        Selection.AutoFilter Field:=Selection.Column, Criteria1:="=#N/A"
    End If
End Sub
