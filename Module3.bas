Sub check_page()

Dim cover_count As Integer
Dim report_count As Integer
Dim total_worksheet_count As Integer
Dim total_page_count As Integer
Dim Idx As Integer

'Remind user to arrange worksheet in page order
MsgBox "PLEASE ENSURE THAT WORKSHEETS ARE ARRANGED IN DESIRED PAGE ORDER." & _
    vbCrLf & vbCrLf & "We assume worksheet Cover Page is first, followed by reports"

'Init some counters
cover_count = 0
report_count = 0
total_worksheet_count = Sheets.count

'Init for appending to any array like object in VBA
Dim page_names() As Variant
Dim page_idx As Integer

'Get no. of pages that starts with "Cover", "report" (i.e. report1, report2)
'Also get names of each sheet
'Assume worksheets are already arranged in page order
page_idx = 0
For Idx = 1 To total_worksheet_count

    If InStr(1, Sheets(Idx).name, "Cover") <> 0 Then
        cover_count = cover_count + 1
        ReDim Preserve page_names(0 To page_idx)
        page_names(page_idx) = Sheets(Idx).name
        page_idx = page_idx + 1
    End If
    
    If InStr(1, Sheets(Idx).name, "report") <> 0 Then
        If InStr(1, Sheets(Idx).name, "mA") = 0 Then
            report_count = report_count + 1
            ReDim Preserve page_names(0 To page_idx)
            page_names(page_idx) = Sheets(Idx).name
            page_idx = page_idx + 1
        End If
    End If

Next Idx

'Sum up for total page count
total_page_count = cover_count + report_count

'Then we go to each worksheet
'Find cell that has string "Page"
'Move left until we encounter a cell which has "of" in it
'Then we set the correct page

Dim name As Variant
Dim current_page As Integer

current_page = 1
For Each name In page_names

    'Get address of cell with "Page" in it
    addr = Worksheets(name).UsedRange.Find("Page", MatchCase:=True, LookAt:=xlWhole).Address
    
    'Move left of cell with "Page" until we encounter a cell that has "of" in it
    'Then we edit page of total pages
    Idx = 1
    Do While True
        If InStr(1, Worksheets(name).Range(addr).Offset(rowOffset:=0, columnOffset:=Idx), "of") <> 0 Then
            Worksheets(name).Range(addr).Offset(rowOffset:=0, columnOffset:=Idx) = current_page & " " & "of" & " " & total_page_count
            Exit Do
        Else
            Idx = Idx + 1
            If Idx > 5 Then
                MsgBox "checking pages: 'of' not found for in worksheet: " & name
                Exit Do
            End If
        End If
    Loop
    
    current_page = current_page + 1
    
Next

MsgBox "Completed checking page of total pages."

End Sub


