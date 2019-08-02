Sub FormatTopRow()
'Purpose: Freezes and formats the top row of the table to make it easier to look at and work with
'Current sheet only
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
