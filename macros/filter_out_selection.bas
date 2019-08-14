Sub FilterOutSelection()
'Purpose: Filter OUT (remove) selected cell from the selection
'Can be used multiple times multiple columns
'Only filters one cell so select first cell cell if multiple are selected
If Selection.Count > 1 Then ActiveCell.Select
'Check for existing filter
If ActiveSheet.AutoFilterMode = False Then Selection.AutoFilter
'Autofilter uses column number relative to the table
filtercolumn = ActiveCell.Column - ActiveSheet.AutoFilter.Range.Column + 1
'Check for error cell
If IsError(Selection.Value) Then cellvalue = Selection.Text Else cellvalue = Selection.Value
'Filter
Selection.AutoFilter Field:=filtercolumn, Criteria1:="<>" & cellvalue, Operator:=xlAnd
End Sub
