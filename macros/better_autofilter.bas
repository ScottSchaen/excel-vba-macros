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
