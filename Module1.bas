Attribute VB_Name = "Module1"
Sub check_date()

'Variables for Data Calibrated
Dim dc_addr As String
Dim date_calibrated As Date

'Variables for Due Date
Dim dd_addr As String
Dim due_dates() As Variant

'Keep track of number of bad dates
Dim bad_date_count As Integer

'Counter for loop control
Dim Idx As Integer

'''''''''''''''''''''''''''''''''''''''''''
Dim dd_str As String
Dim dc_str As String
Dim sheet_name As String
Dim no_of_rows_below_dueDate As Integer
Dim no_of_cols_rightof_dateCal As Integer

'PARAMETERS

'Search strings to know where to start finding the dates we want.
'Asterisk here to cater for spaces and what nots
dc_str = "Date Calibrated*"
dd_str = "Due Date*"

'Name of sheet that we want to perform check_date
sheet_name = "Cover Page"

'E.g: If the first date is right below col Due Date, i.e. no extra row in betw, we set to 1
no_of_rows_below_dueDate = 1

'E.g: If date calibrated is 3 cells right of "Date Calibrated",  we set to 3
no_of_cols_rightof_dateCal = 3

'''''''''''''''''''''''''''''''''''''''''''

'Find dc_str, move 3 cells left and get date
dc_addr = Worksheets(sheet_name).UsedRange.Find(dc_str, MatchCase:=True).Address
date_calibrated = Worksheets(sheet_name).Range(dc_addr).Offset(rowOffset:=0, columnOffset:=no_of_cols_rightof_dateCal)

'Find dd_str, move 1 cell down, get all dates until empty cell
dd_addr = Worksheets(sheet_name).UsedRange.Find(dd_str, MatchCase:=True).Address

'Get all due dates and store it to an array
Idx = 0
While Not IsEmpty(Worksheets(sheet_name).Range(dd_addr).Offset(rowOffset:=Idx + no_of_rows_below_dueDate, columnOffset:=0))
    ReDim Preserve due_dates(0 To Idx)
    due_dates(Idx) = Worksheets(sheet_name).Range(dd_addr).Offset(rowOffset:=Idx + no_of_rows_below_dueDate, columnOffset:=0)
    Idx = Idx + 1
Wend

'Compare dates: if dc_str is later than any dates in dd_str then pop msgbox to warn
'else pop msgBox to say all is ok
bad_date_count = 0
For Each date_to_check In due_dates
    If DateDiff("d", date_calibrated, date_to_check) < 0 Then
        MsgBox ("Ref Instrument Due Date: " & date_to_check & " is OLDER than Date Calibrated: " & date_calibrated)
        bad_date_count = bad_date_count + 1
    End If
Next date_to_check

If bad_date_count = 0 Then
    MsgBox ("All Date Check Passed!")
End If

End Sub
