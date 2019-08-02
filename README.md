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
With a single click this macro will select all cells with formulas on the page. This is useful if you’re going to publish or share a spreadsheet and want the values hard coded.

## #N/A Check
Don’t be the guy or gal that sends out spreadsheets with `#N/A` all over it. Use this macro to highlight all of these in your current tab. You can prevent `#N/A` by wrapping your formula in an `iferror(old_formula,displayed_value_if_error)`.

## Filter for ONLY Selected
I was using this macro before it was built into Excel. It will filter your table and show you just values of the cell you have selected. Alternatively, you can right click on the cell and go to `Filter` → `Filter by Selected Cell’s Value`

## Filter out (remove) Selected
This does the opposite of above and filters out or removes only the selected value from your table. For instance, say you have a list of orders and want to remove all orders with a $0 value. Just click $0 in the table and then run this macro. 

## Top Left Active Cell
This is a great feature if you share spreadsheets with a lot of tabs. It simply cycles through all your sheets placing the active cell on the top left.

## Remove External Links
If you’re sharing spreadsheets and you occasionally reference other workbooks, this macro is a must. It will scan through your workbook looking for external formulas and then will replace the link with the value.

## Select Uniques
This can be achieved a few ways in Excel, but I like my way best :) It selects only unique values in your selection. There’s a number of use-cases here. 

## Comma Separate Selection
This is a really useful feature if you use SQL or use a BI tool that accepts comma separated values. It simply takes all of your cells in a selection and comma separates them into a near by cell. It can be used with the `Select Uniques` macro to only comma separate unique values in a selection.

## Macro Caveats:
* There’s no undo!
* You gotta get personal.xls working
* You want the right icon for your macro, but you’re limited
* Some don’t work on Mac
