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
