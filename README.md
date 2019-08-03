# Excel Shortcuts - Helpful Macros in VBA
Here are some macros I've incorporated into my day-to-day Microsoft Excel usage over the years. They help cut down on clicks so you can be more efficient with time.
For best results, you'll want to save these so they are always available and set them up as commands/buttons in your `HOME` ribbon.

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
    
## Number Format
I usually want my numbers formatted like this: `452,199`  
Not like this: `452199`  
Not like this: `452,199.00`  
And not like this: `|`&nbsp;&nbsp;&nbsp;&nbsp;`452,199|`

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
With a single click this macro will select all cells containing a formula on the active sheet. This is useful if you’re going to publish or share a spreadsheet and want the values hard coded.

```bas
On Error GoTo err
    Cells.SpecialCells(xlCellTypeFormulas).Select
Exit Sub
err:
    If err.Number = 1004 Then MsgBox "No Formulas Here!"
```

## #N/A Check
Don’t be the guy or gal that sends out spreadsheets with `#N/A` all over it. Use this macro to highlight all of these in your current tab. You can prevent `#N/A` by wrapping your formula in an `iferror(your_formula,value_if_error)`.

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
'Only works with one cell selected
On Error GoTo err
    'Try filtering to selected
    Selection.AutoFilter Field:=Selection.Column, Criteria1:=Selection.Value
Exit Sub
err:
    If err = 1004 Then
        'Turn on autofilter if it's not on already
        Selection.AutoFilter
        Selection.AutoFilter Field:=Selection.Column, Criteria1:=Selection.Value
    Else
        'If it doesn't work, filter to '#N/A'
        Selection.AutoFilter Field:=Selection.Column, Criteria1:="#N/A"
    End If
```

## Filter out (remove) Selected
This does the opposite of above and filters out or removes only the selected value from your table. For instance, say you have a list of orders and want to remove all orders with a $0 value. Just click $0 in the table and then run this macro. 

```bas
'Only works with one cell selected
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
If you’re sharing spreadsheets and you occasionally reference other workbooks, this macro is a must. It will scan through your workbook looking for external formulas and then will replace the link with the value.

## Select Uniques
This can be achieved a few ways in Excel, but I like my way best :) It selects only unique values in your selection. There’s a number of use-cases here. 

```bas
'Selection does not need to be a single range, but it does need to be on the same sheet.
If Selection.Count > 5000 Then
    response = MsgBox("This could take a while", vbOKOnly + vbInformation)
    If response = vbCancel Then Exit Sub
End If
Dim d As Dictionary
Set d = New Dictionary
Dim uniques As Range
For Each cell In Selection.Cells
    If first = False Then Set uniques = cell
    first = True
    'Removes blank cells from consideration
    If d(cell.Value) = False And cell.Value <> "" Then
        d(cell.Value) = True
        Set uniques = Union(uniques, cell)
    End If
Next cell
uniques.Select
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

## Macro Caveats:
* There’s no undo!
* You gotta get personal.xls working
* You want the right icon for your macro, but you’re limited
* Some don’t work on Mac
