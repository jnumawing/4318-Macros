Attribute VB_Name = "Module1"
Sub opening_zoom()
'delcaring the variables
    Dim zoom  As Variant
    Dim cell_range As Range
    Dim location  As String
    
    'using Application.GetOpenFilename method to filter the csv and thus open the file
    zoom = Application.GetOpenFilename("CSV File (*.csv), *.csv")
    
    'asking the user to input the A columb to place the cvs information
    Set cell_range = Application.InputBox("Cell to begin copying this CSV File: ", "Select Range", Type:=8)
  
  ' assigning the location to range adresss
    location = cell_range.Address
    
    'Query Table
    ' refrence
    ' https://docs.microsoft.com/en-us/office/vba/api/excel.querytables.add
    ' addign the information into the excel sheet based on selection use only Columb A becuase I cleared the other columbs
    With ActiveSheet.QueryTables.Add("TEXT;" & zoom, Range(location))
        ' same names and rows
        .FieldNames = True
        .RowNumbers = False
        ' not not have the same formating
        .PreserveFormatting = False
        .RefreshOnFileOpen = False
        ' adds the needed rows
        .RefreshStyle = xlInsertDeleteCells
        ' sabe data in the worksheet
        .SaveData = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 936
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileSpaceDelimiter = False
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    
    
  '  Here I edit the contents of the cvs file to remove unwated information
    Columns("B:B").Select
    Selection.ClearContents
    Columns("C:C").Select
    Selection.ClearContents
    Columns("D:D").Select
    Selection.ClearContents
    Columns("E:E").Select
    Selection.ClearContents
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "User Name "
    Range("B1").Select
    
'  calling the sub fucntion to find the doffrence bewteen the two sheets
   Call find_diffrences("sheet1", "sheet2")
End Sub


' to find the difrences between the two sheets
Sub find_diffrences(zoom_values As String, orginal_values As String)

' declare variables
Dim sheet As Range
Dim answer As Integer
Dim count As Integer

' literate thorught the cell worksheets
For Each sheet In ActiveWorkbook.Worksheets(orginal_values).UsedRange

' if the value is not the same highlight it red
    If Not sheet.Value = ActiveWorkbook.Worksheets(zoom_values).Cells(sheet.Row, sheet.Column).Value Then
        ' highling the different cells on the sheet #2
        sheet.Interior.Color = vbRed
        ' count to display how manny diffrence were found
        count = count + 1
        
    End If
Next

If count = 0 Then
' display hwo many students were not there at the zoom meeting
MsgBox " All the students were at the end of the meeting", vbOKOnly
Else
MsgBox count & " students were NOT at the end of the meeting", vbOKOnly
End If

ActiveWorkbook.Sheets(orginal_values).Select

' ask the user if he wants to clear the contents of the zoom values
answer = MsgBox("Do you want to clear contents of sheet 1?", vbQuestion + vbYesNo)
' if yes then clear contents
If answer = vbYes Then
Sheets("sheet1").Cells.Clear
Else
    MsgBox "okay"
  End If
End Sub


