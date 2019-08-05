'Remove Attribute VB_Name line if copying and pasting into a VBA Module. Keep if downloading file and importing.
Attribute VB_Name = "ScottsMacros"

Sub FormatTopRow()
'Purpose: Freezes and formats the top row of the table to make it easier to look at and work with
'Active sheet only
Dim toprow As Range
'Header column needs to be on row 1, but this can be changed.
'Looks for right-most used column
Set toprow = Range("A1:" & Range("IV1").End(xlToLeft).Address)
Range("A2").Select
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


Sub NumberFormat()
'Purpose: Formats numbers by adding comma, removing decimals, and centering
Selection.NumberFormat = "#,##0"
Selection.HorizontalAlignment = xlCenter
End Sub


Sub BetterAutoFilter()
'Purpose: One button that turns on autofilter (when off), clear the filter (when filtered), or shut autofilter (when on and not filtered)
'Requires more buttons and clicks otherwise
On Error Resume Next
If ActiveSheet.FilterMode = True Then
    ActiveSheet.ShowAllData
Else
    Selection.AutoFilter
End If
End Sub


Sub CheckForFormulas()
'Purpose: Check for formulas in current sheet
On Error GoTo err
    Cells.SpecialCells(xlCellTypeFormulas).Select
Exit Sub
err:
    If err.Number = 1004 Then MsgBox "No Formulas Here!"
End Sub


Sub CheckForNAs()
'Purpose: Check for #N/As in current sheet
On Error GoTo err
    Cells.SpecialCells(xlCellTypeFormulas, xlErrors).Select
Exit Sub
err:
    If err.Number = 1004 Then MsgBox "No Errors Here!"
End Sub


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


Sub TopLeftActiveCell()
'Purpose: Sets active cell to top left ($A$1) for all sheets
Dim currsheet As Worksheet
Dim sheet As Worksheet
Set currsheet = ActiveSheet
'Change A1 to suit your preference
Const TopLeft As String = "A1"
'Loop through all the sheets in the workbook
For Each sheet In Worksheets
    'Only does this for visible worksheets
    If sheet.Visible = xlSheetVisible Then Application.GoTo sheet.Range(TopLeft), Scroll:=True
Next sheet
currsheet.Activate
End Sub


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
    'For non-blank cells...
    If cell.Value <> "" Then
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
