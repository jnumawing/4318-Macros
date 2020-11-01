Sub ID()
'
' ID Macro
' Outputs last 4 digits of IDs
' Checks if there are any repeats
' Will change output to last 5 digits if no repeats
' Note: Original IDs must be left of selected region
' Ex: If Original IDs are in A, output will be in B
' Therefore highlight columns right of original inputs
' Numbers must be formatted similar to Bronco ID
'
    Dim myRange As Range
    Dim myCell As Range
    Dim rep As Boolean
    Set myRange = Selection
    '4 digit ID
    For Each myCell In myRange
        myCell.Value = "=MID(RC[-1],6,4)"
    Next myCell
    
    For Each myCell In myRange
        If WorksheetFunction.CountIf(myRange, myCell.Value) > 1 Then
        rep = True
    End If
    Next myCell
    '5 digit ID if there are repeats
    For Each myCell In myRange
        If (rep) Then
        myCell.Value = "=MID(RC[-1],5,5)"
    End If
    Next myCell
    '6 digit ID if there are still repeats
    'last case
    'can add more if there are still repeats
    'by following patterns
    For Each myCell In myRange
        If (rep) Then
        myCell.Value = "=MID(RC[-1],4,6)"
    End If
    Next myCell
End Sub

'this macro is used to format IDs
'Bronco IDs are 9 digit
'if IDs have zeros at the front, macro will add zeros even if
'excel automatically removes them
'select entered values to format
Sub ID_Format()
  'used to format inputs
  On Error Resume Next
  Dim nRange As Range
  Set nSet = nRange
  nRange.Select
  Selection.SpecialCells(xlCellTypeConstants, 1).Select
  Selection.NumberFormat = "00000000#"

  nRange.Select
End Sub
