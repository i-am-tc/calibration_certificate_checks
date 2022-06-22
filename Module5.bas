Attribute VB_Name = "Module5"
Function criteria_value() As Double

Dim cr_addr As String
Dim criteria_accur As Variant
Dim crieria_range As Variant
Dim no_of_cols_right_accur As Integer
Dim no_of_cols_right_range As Integer
Dim sheet_name_criteria As String

'''''''''''''''''''''''''''''''''''''''''''

'Parameters

sheet_name_criteria = "Cover Page"

'No of cols to move right to read accuracy & range values
no_of_cols_right_accur = 3
no_of_cols_right_range = 3

'''''''''''''''''''''''''''''''''''''''''''

'Get criteria value

'Get cell address of Accuracy
cr_addr = Worksheets(sheet_name_criteria).UsedRange.Find("Accuracy*", MatchCase:=True).Address
'Get cell's value for Accuracy, with offset
criteria_accur = Split(Worksheets(sheet_name_criteria).Range(cr_addr).Offset(rowOffset:=0, columnOffset:=no_of_cols_right_accur), " ")(0)

'Sanity check if % in criteria_accur
If InStr(1, criteria_accur, "%") Then
    'Remove % symbol
    criteria_accur = Replace(criteria_accur, "%", "")
End If
'Convert to double
criteria_accur = CDbl(criteria_accur)

'Get cell address of Range
cr_addr = Worksheets(sheet_name_criteria).UsedRange.Find("Range*", MatchCase:=True).Address
'Get cell's value for Range, with offset
criteria_range = Worksheets(sheet_name_criteria).Range(cr_addr).Offset(rowOffset:=0, columnOffset:=no_of_cols_right_range)

'check if 'to' in criteria_range
If InStr(1, criteria_range, "to") Then
    'Get number to left of to
    left_val = CDbl(Split(criteria_range, " ")(0))
    'Get number to right of to
    right_val = CDbl(Split(criteria_range, " ")(2))
    'Sum to give criteria_range
    criteria_range = Abs(left_val) + Abs(right_val)
'check if '+/-' in criteria_range
ElseIf InStr(1, criteria_range, "+/-") Then
    'Get number to right of +/-
    criteria_range = CDbl(Split(criteria_range, " ")(1)) * 2
Else
    criteria_range = CDbl(Split(criteria_range, " ")(0))
End If


'Final compute for criteria_value
criteria_value = ((criteria_accur / 100) * criteria_range)

End Function
