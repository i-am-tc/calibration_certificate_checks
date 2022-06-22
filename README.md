# calibration_cetificate_checks

A series of simple checks written in VBA to automate mundane checks for a calibration certificate.

dev 001 : A sample calibration certificate macro enabled Excel. Each macro explained below is assinged to a button in the corresponding worksheet that it is supposed to operate on.

Module1 , check_date: worksheet "Cover Page", compare "Date Calibrated" in cell L10 against dates under "Reference Instruments, Due Dates", cells L26 and below. If "Due Dates" are older than "Date Calibrated", we throw out a msg box to tell operator.

Module2, check_error: worksheet "report1", for each column titled "Error", compare against a criteria value (calculated in Module5). If it is larger than, then we add an asterisk in columns E or I

Module3, check_page: worksheet "Cover Page", consider "Cover Page" as 1st page, we check how many worksheets has the substring "report" in them. Then we count how many pages we have in total, go into each worksheet "report" and edit the pagination (PAGE of TOTAL_PAGE ) to the correct one.

Module4, check_low_high: worksheet "Calibration Data", fill out columns "Low" and "High" (Acceptance Limits) according to criteria value calculated in Module5.

Module5, criteria_value : Get accuracy and range from "Cover Page", account for various cases of range inputs, calculate criteria value for comparison in other sub routines.
