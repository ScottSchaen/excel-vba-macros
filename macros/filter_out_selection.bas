Sub FilterOutSelection()
'Purpose: Filter OUT (remove) selected cell from the selection
'Only filters one cell so this reduces the selection to one cell if multiple are selected
If Selection.Count > 1 Then ActiveCell.Select
On Error GoTo err
    Selection.AutoFilter Field:=Selection.Column, Criteria1:="<>" & Selection.Value, Operator:=xlAnd
Exit Sub
err:
    If err = 1004 Then
        'Turn on autofilter if it's not on already
        Selection.AutoFilter
        Selection.AutoFilter Field:=Selection.Column, Criteria1:="<>" & Selection.Value, Operator:=xlAnd
    Else
        'If it doesn't work, filter out '#N/A'
        Selection.AutoFilter Field:=Selection.Column, Criteria1:="<>#N/A", Operator:=xlAnd
    End If
End Sub
