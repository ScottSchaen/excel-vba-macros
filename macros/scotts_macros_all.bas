'Remove Attribute VB_Name line if copying and pasting into a VBA Module. Keep if downloading file and importing.
Attribute VB_Name = "ScottsMacros"

'SCOTTS MACROS
'http://linkedin.com/in/ScottSchaen

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
'Can be used multiple times on multiple columns
'Only filters one cell so select first cell cell if multiple are selected
If Selection.Count > 1 Then ActiveCell.Select
'Check for existing filter
If ActiveSheet.AutoFilterMode = False Then Selection.AutoFilter
'Autofilter uses column number relative to the table
filtercolumn = ActiveCell.Column - ActiveSheet.AutoFilter.Range.Column + 1
'Check for error cell
If IsError(Selection.Value) Then cellvalue = Selection.Text Else cellvalue = Selection.Value
'Filter
Selection.AutoFilter Field:=filtercolumn, Criteria1:="=" & cellvalue
End Sub


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
