Sub Current_Date() 'Selected cells will display the current date
Dim cell As Range 'Set cell as Range datatype
Selection.Value = Selection.Value
For Each cell In Selection 'for all cells selected
cell = Date 'make the cell current date
Next cell
End Sub

Sub Gen_Rand_Dates() 'Generated random dates from 1/1/1900 to todays date (useful for generating lots of data quickly
Dim start As Date 'set start and enddate as Date datatype
Dim enddate As Date
start = "1/1/1900" 'starting date is 1/1/1900
enddate = Date 'end date is current date
Dim cell As Range '
Selection.Value = Selection.Value
For Each cell In Selection 'for each cell selected, display random date between start and end dates
cell = WorksheetFunction.RandBetween(start, enddate)
cell = Format(cell, "m/d/yyyy") 'format the generated dates
Next cell
End Sub

Sub To_Day() 'detects if a cell contains a date and replaces it with just the day
Dim dates As Range
Selection.Value = Selection.Value
For Each dates In Selection 'for each date that is selected
If IsDate(dates) = True Then 'if the cell contains a date
With dates
.Value = Day(dates) 'set the value to the day of the date
.NumberFormat = "General" 'the format of the date is the general format
End With
End If
Next dates
End Sub

Sub To_Month() 'detects if a cell contains a date and replaces it with just the month
Dim dates As Range
Selection.Value = Selection.Value
For Each dates In Selection 'for each date that is selected
If IsDate(dates) = True Then 'if the cell contains a date
With dates
.Value = Month(dates) 'set the value to the month of the date
.NumberFormat = "General" 'the format of the date is the general format
End With
End If
Next dates
End Sub

Sub To_Year() 'detects if a cell contains a date and replaces it with just the year
Dim dates As Range
Selection.Value = Selection.Value
For Each dates In Selection 'for each date that is selected
If IsDate(dates) = True Then 'if the cell contains a date
With dates
.Value = Year(dates) 'set the value to the year of the date
.NumberFormat = "General" 'the format of the date is the general format
End With
End If
Next dates
End Sub
