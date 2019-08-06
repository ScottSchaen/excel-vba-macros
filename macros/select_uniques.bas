Sub SelectUnique()
'Purpose: Select only unique values in selection. Effectively removes duplicates from selection.
'Selection does not need to be a single range, but it does need to be on the same sheet.
If Selection.Count > 5000 Then
    response = MsgBox("This could take a while", vbOKCancel + vbInformation)
    If response = vbCancel Then Exit Sub
End If
ReDim vals(Selection.Count)
Dim uniques As Range
'Cycle through all values in selection
For Each cell In Selection
    'Skip blank cells and errored cells
    If Not IsError(cell) And Not IsEmpty(cell) Then
        'Set first value
        If uniques Is Nothing Then
            Set uniques = cell
            vals(1) = cell.Value
            uniq_counter = 2
        End If
        'Check each cell against previously set unique values
        For checker = 1 To uniq_counter - 1
            If vals(checker) = cell.Value Then Exit For
            If checker = uniq_counter - 1 Then
                Set uniques = Union(uniques, cell)
                vals(uniq_counter) = cell.Value
                uniq_counter = uniq_counter + 1
            End If
        Next checker
    End If
Next cell
'Select unique range if it exists
If Not uniques Is Nothing Then uniques.Select
End Sub
