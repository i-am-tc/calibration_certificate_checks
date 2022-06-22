Attribute VB_Name = "Module2"
Sub check_error()

Dim err_addrs(0 To 1) As Variant
Dim criteria As Double
Dim Idx As Integer
Dim empty_count As Integer
Dim err_first_col() As Variant
Dim err_second_col() As Variant
Dim error_val As Range
Dim have_error_flag As Boolean

'''''''''''''''''''''''''''''''''''''''''''
Dim err_col_name As String
Dim sheet_name_err_chk As String
Dim sheet_name_criteria As String
Dim no_of_rows_below_error As Integer
Dim asterisk_col_offset As Integer

'Parameters

'Col name that has our error values
err_col_name = "Error"
'Where to perform the error checks
sheet_name_err_chk = ActiveSheet.name
'Where to find criteria for compare
sheet_name_criteria = "Cover Page"
'E.g: If the first error val is 2 rows below col Error, we set to 2
no_of_rows_below_error = 2

'Since there is 1 empty row (i.e row 24), in between for Error col in report (print),
'we keep checking after 1st empty cell
'and only stop check when we encounter the 2nd empty cell.
'use this if structure of worksheet changes.
empty_allow = 2

asterisk_col_offset = -3

'''''''''''''''''''''''''''''''''''''''''''

'We expect 2 cells with err_col_name.
err_addrs(0) = Worksheets(sheet_name_err_chk).UsedRange.Find(err_col_name, MatchCase:=True).Address
err_addrs(1) = Worksheets(sheet_name_err_chk).UsedRange.FindNext(Range(err_addrs(0))).Address

'Get criteria value
criteria = Module5.criteria_value()

'Go to each cell of 1st error column, check against criteria
'If fail, flag abnormal colors. if pass, flag normal colors
have_error_flag = False
For Each addr In err_addrs
    Idx = 0
    empty_count = 0
    Do While True
        
        'Set objects for error value and measurement value.
        Set err_val = Worksheets(sheet_name_err_chk).Range(addr).Offset(rowOffset:=Idx + no_of_rows_below_error, columnOffset:=0)
        
        'Check if we have reached the end
        If IsEmpty(err_val) Then
            empty_count = empty_count + 1
        End If
        
        If empty_count >= empty_allow Then
            Exit Do
        End If
        
        'If criteria not met, we enter asterisk, in 2 cols left. Else, we enter blank
        If Abs(err_val) > criteria Then
            have_error_flag = True
            err_val.Offset(rowOffset:=0, columnOffset:=asterisk_col_offset) = "*"
            err_val.Offset(rowOffset:=0, columnOffset:=asterisk_col_offset).HorizontalAlignment = xlCenter
            err_val.Offset(rowOffset:=0, columnOffset:=asterisk_col_offset).Font.Size = 16
        Else
            err_val.Offset(rowOffset:=0, columnOffset:=asterisk_col_offset) = " "
        End If
            
        Idx = Idx + 1
        
    Loop
Next

If have_error_flag = True Then
    MsgBox "Errors have been flagged with asterisks. Please check."
Else
    MsgBox "No Errors have been flagged."
End If


End Sub

