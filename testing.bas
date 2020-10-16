Sub dates()
Selection.Value = Selection.Value
If Selection.Value = "date" Then
    Selection.Value = Date
End If
End Sub

Sub To_Day()
Dim dates As Range
Selection.Value = Selection.Value
For Each dates In Selection
If IsDate(dates) = True Then
With dates
.Value = Day(dates)
.NumberFormat = "General"
End With
End If
Next dates
End Sub
