Sub KillExternalFormulas()
'Purpose: Replaces external formulas that link to other workbooks with their values
'Only applies to cells in selection
Dim replaced As Integer
replaced = 0
wholebook = MsgBox("Do you want to Remove External Formulas from the whole WORKBOOK? Click no for active sheet or selection. You Can't Undo This!!", vbYesNoCancel + vbInformation, "Apply to whole WORKBOOK?")
If wholebook = vbCancel Then Exit Sub
If wholebook = vbNo Then
    wholesheet = MsgBox("Do you want to Remove External Formulas from the whole WORKSHEET? Click no if you just want to remove from the selection. You Can't Undo This!!", vbYesNoCancel, "Apply to whole WORKSHEET?")
    If wholesheet = vbCancel Then Exit Sub
    If wholesheet = vbYes Then ActiveSheet.UsedRange.Select
    For Each cell In Selection
        If InStr(cell.Formula, "!") > 0 Then
            cell.Value = cell.Value
            replaced = replaced + 1
        End If
    Next cell
Else
    For Each sheet In ActiveWorkbook.Worksheets
        For Each cell In sheet.UsedRange
            If InStr(cell.Formula, "!") > 0 Then
               cell.Value = cell.Value
               replaced = replaced + 1
            End If
        Next cell
    Next sheet
End If
MsgBox replaced & " formula(s) removed!"
        
End Sub
