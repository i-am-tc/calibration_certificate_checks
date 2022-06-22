Attribute VB_Name = "Module4"
Sub check_low_and_high()

Dim low_addr As String
Dim high_addr As String
Dim testdata_addr As String
Dim sheet_name_lowhigh As String
Dim row_2nd_testdata As String
Dim Idx As Integer

'''''''''''''''''''''''''''''''''''''''''''

'Parameters

sheet_name_lowhigh = ActiveSheet.name
cell_to_search_after = "$A25"
reference_col_name = "Test Data*"

'''''''''''''''''''''''''''''''''''''''''''

'Find cell addresses of low, high and test data cols.
'We may have 2 test data cols. Hence, we use After with .Find to get the one we want
low_addr = Worksheets(sheet_name_lowhigh).UsedRange.Find("Low*", MatchCase:=True).Address
high_addr = Worksheets(sheet_name_lowhigh).UsedRange.Find("High*", MatchCase:=True).Address
testdata_addr = Worksheets(sheet_name_lowhigh).UsedRange _
    .Find(reference_col_name, MatchCase:=True, After:=Range(cell_to_search_after)).Address
    
'Get criteria value
criteria = Module5.criteria_value

'Go to each cell of test data cols
'Add criteria and put it to high col
'Subtract criteria and put it to low col
Idx = 1
Do While True
    
    'Set objects for error value and measurement value.
    Set test_val = Worksheets(sheet_name_lowhigh).Range(testdata_addr).Offset(rowOffset:=Idx, columnOffset:=0)
    
    'Check if we have reached the end
    If IsEmpty(test_val) Then
        Exit Do
    End If
    
    'Add criteria to test data and deposit it to high col
    Worksheets(sheet_name_lowhigh).Range(high_addr).Offset(rowOffset:=Idx, columnOffset:=0) = test_val + criteria
        
    'Subtract criteria to test data and deposit it to low col
    Worksheets(sheet_name_lowhigh).Range(low_addr).Offset(rowOffset:=Idx, columnOffset:=0) = test_val - criteria
        
    Idx = Idx + 1
    
Loop

MsgBox "Completed filling out of LOW and HIGH column"

End Sub
