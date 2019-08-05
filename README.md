# VBA Macros to Super Charge Microsoft Excel
Become a Microsoft Excel power user with these handy VBA macros. I've been using and fine-tuning these for years to help make my day-to-day tasks more efficient. Take them and make them your own. You can download all the macros at (macros/scotts_macros_all.bas)[/macros/scotts_macros_all.bas]. If you're new to VBA and Excel macros you'll want to read my notes on [getting started](#how-to-get-started). Be sure to add these as commands/buttons to Excel's `HOME` ribbon to really make them useful!

# Contents
* [*How to get started*](#how-to-get-started)
* [*Notes and Caveats*](#macro-notes--caveats)
* [Format Top Row of your table](#format-top-row)
* [Better number format](#better-number-format)
* [Better AutoFilter](#better-autofilter)
* [Check worksheet for formulas](#formula-check)
* [Check worksheet for #N/As](#na-check)
* [Filter table for selected cell](#filter-for-only-selected)
* [Filter table and remove selected cell](#filter-out-remove-selected)
* [Reset active cell to top left for all sheets in workbook](#reset-active-cell-to-top-left-for-all-sheets-in-workbook)
* [Remove external links](#remove-external-links)
* [Select Uniques (by removing duplicates from selection)](#select-uniques)
* [Comma Separate Selection](#comma-separate-selection)



## Format Top Row
This may be my most-used macro. In one click it format the table header and freeze the top pane. It makes tables a lot easier on the eyes and knows exactly what to format.

```bas
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
```
[➥full code](/macros/format_top_row.bas)
    
## Better Number Format
I usually want my numbers formatted like this: `452,199`  
Not like this: `452199`  
Not like this: `452,199.00`  
And not like this: `|_____452,199|` (right justified)

This means centered, with a comma separator, and no decimals. Crazily, the only way to do this is with many clicks (I think 8 is the least) through the `Format Cells` dialog. FTFY:

```bas
Selection.NumberFormat = "#,##0"
Selection.HorizontalAlignment = xlCenter
```

## Better AutoFilter
I filter my tables a lot, so I made one button that enables auto-filter on a table, clears any existing filters, and shuts auto-filter. It cuts down on clicks and is really how the auto-filter button should work.

```bas
On Error Resume Next
If ActiveSheet.FilterMode = True Then
    ActiveSheet.ShowAllData
Else
    Selection.AutoFilter
End If
```

## Formula Check
With a single click this macro will select all cells containing a formula on the active sheet. This is useful if you’re going to publish or share a spreadsheet and want the values hard coded. After running it, you can look at the status bar (bottom right) to see how many cells/formulas are selected.

```bas
On Error GoTo err
    Cells.SpecialCells(xlCellTypeFormulas).Select
Exit Sub
err:
    If err.Number = 1004 Then MsgBox "No Formulas Here!"
```

## #N/A Check
Don’t be the guy or gal that sends out spreadsheets with `#N/A` all over it. Use this macro to highlight all of these in your current tab. It will catch other types of error cells too, like `DIV/0!`. You can prevent `#N/A` by wrapping your formula in an `iferror(your_formula,value_if_error)`. After running it, you can look at the status bar (bottom right) to see how many cells/#NAs are selected.

```bas
On Error GoTo err
    Cells.SpecialCells(xlCellTypeFormulas, xlErrors).Select
Exit Sub
err:
    If err.Number = 1004 Then MsgBox "No Errors Here!"
```

## Filter for ONLY Selected
I was using this macro before it was built into Excel. It will filter your table and show you just values of the cell you have selected. Alternatively, you can right click on the cell and go to `Filter` → `Filter by Selected Cell’s Value`

```bas
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
```

## Filter out (remove) Selected
This does the opposite of above and filters out or removes only the selected value from your table. For instance, say you have a list of orders and want to remove all orders with a $0 value. Just click $0 in the table and then run this macro. 

```bas
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
```

## Reset active cell to top left for all sheets in workbook
This is a great feature if you share spreadsheets with a lot of tabs. It simply cycles through all your sheets placing the active cell on the top left. Useful when you're sharing spreadsheets

```bas
Dim currsheet As Worksheet
Dim sheet As Worksheet
Set currsheet = ActiveSheet
'Change A1 to suit your preference
Const TopLeft As String = "A1"
'Loop through all the sheets in the workbook
For Each sheet In Worksheets
    'Only does this for visible worksheets
    If sheet.Visible = xlSheetVisible Then Application.GoTo sheet.Range(TopLeft), scroll:=True
Next sheet
currsheet.Activate
```

## Remove External Links
If you’re sharing spreadsheets and you occasionally reference other workbooks, this macro is a must. The macro gives you a few options for replacing external references with their values -- you can just remove external references in the selected cells, or the entire active worksheet, or the entire workbook. You can't undo this function so use with caution!

```bas
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
```

## Select Uniques
This can be achieved a few ways in Excel, but I like my way best :) It selects only unique values in your selection. There’s a number of use-cases here. 

```bas
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
```

## Comma Separate Selection
This is a really useful feature if you use SQL or use a BI tool that filters on comma separated values. It simply takes all of your cells in a selection and comma separates them into a near by cell. The macro will ask you if you want to wrap the values in quotes (for strings). It can be used with the `Select Uniques` macro to only comma separate unique values in a selection.

```bas
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
```

## Macro Notes & Caveats:
* There’s no undo for a macro! (Unless you program one in)
* You need to get `PERSONAL.XLSB` working so the macros are always available. You also want to add these macros as buttons on your Ribbon. See next section for both.
* Excel for Mac has come a really long way, but you can't currently choose a custom icon for Macros on your ribbon :(
* These macros have been working like a charm for me, but there's always room for improvement.
* If you're new to macros, you can learn a lot from recording yourself doing it and/or googling "VBA + thing-you're-trying-to-do". Recording macros is really useful but try to remove the fluff and absolutely references that it writes.
* "Step Into" your macros to go line by line and see what's happening as it runs. You can drag variables or statements to the "Watch Window" to see how they're evaluated as you step through.
* Try to change, tweak, add to these to make them more personalized for you!
* If you copy and paste from above, be sure to wrap it in `Sub WhateverMacroYouWant()` and `End Sub`.
* Find me on (LinkedIn)[https://www.linkedin.com/in/scottschaen/] and send me some feedback, or propose a file change by forking this project.

## How To Get Started:
### You need to create a "Personal Macro Workbook" so that **your macros are always available** when Excel is open.

You can read the [Windows Documentation](https://support.office.com/en-gb/article/copy-your-macros-to-a-personal-macro-workbook-aa439b90-f836-4381-97f0-6e4c3f5ee566#OfficeVersion=Windows) or the [Mac Documentation](https://support.office.com/en-gb/article/copy-your-macros-to-a-personal-macro-workbook-aa439b90-f836-4381-97f0-6e4c3f5ee566#OfficeVersion=macOS) but the gist is this:  
&nbsp;&nbsp;&nbsp; a) Enable the Developer tab for your Excel ribbon  
&nbsp;&nbsp;&nbsp; b) Click `Record Macro` and choose to store the macro in "Personal Macro Workbook"  
&nbsp;&nbsp;&nbsp; c) `Stop Recording` the macro and click the `Visual Basic` button (or press <kbd>alt</kbd><kbd>F11</kbd>)  
&nbsp;&nbsp;&nbsp; d) On the project explorer (top left) find `PERSONAL.XLSB`, expand `Modules`, and that's where you want to store all of your macros. You can leave them all in `Module1` or separate them. I prefer less modules, but it doesn't make a huge difference. Remember to Save!

### When you have your macros saved in `PERSONAL.XLSB` you want to **customize the ribbon** and add them as commands/buttons there.

**Windows:** Right click anywhere on the ribbon and select `Customize the Ribbon...`
**Mac:** `Excel` → `Preferences` → `Ribbon & Toolbar`

<p align="center">
  <img width="900" src="/images/macro_ribbon_and_config.png">
</p>

(You can read about this in my 5 Stupid Easy Excel Tips)[https://github.com/ScottSchaen/stupid-easy-excel-tips/blob/master/README.md#5-customize-the-home-ribbon--load-it-up-with-only-useful-functions]

**Happy Excelling,**  
**Scott**
