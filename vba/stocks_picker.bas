Attribute VB_Name = "stocks_picker"
' VBA module for the Stock Picker workbook.
'
' Each macro:
'   1. Checks the JobRunning flag (Main!B14) — refuses to launch if another
'      job is already running. Prevents accidental double-clicks from
'      launching two parallel jobs.
'   2. Sets JobRunning = TRUE and resets StopRequested = FALSE.
'   3. Calls Python via xlwings RunPython (synchronous — VBA blocks until
'      Python returns).
'   4. On return, clears JobRunning back to FALSE.
'
' During the Python run, the STOP checkbox writes TRUE to Main!B13 via
' Excel's internal cell engine (Form Controls update LinkedCell instantly
' even while VBA is blocked). Python polls that cell every 25 tickers and
' breaks out cleanly when set.
'
' RunPython comes from the xlwings Excel add-in.

Private Const JOB_RUNNING_CELL As String = "B14"
Private Const STOP_REQUESTED_CELL As String = "B13"

Private Function IsJobRunning() As Boolean
    Dim v As Variant
    v = ThisWorkbook.Sheets("Main").Range(JOB_RUNNING_CELL).Value
    IsJobRunning = (UCase$(CStr(v)) = "TRUE")
End Function

Private Sub SetJobRunning(running As Boolean)
    If running Then
        ThisWorkbook.Sheets("Main").Range(JOB_RUNNING_CELL).Value = "TRUE"
    Else
        ThisWorkbook.Sheets("Main").Range(JOB_RUNNING_CELL).Value = "FALSE"
    End If
End Sub

Private Sub ClearStopRequest()
    ThisWorkbook.Sheets("Main").Range(STOP_REQUESTED_CELL).Value = "FALSE"
End Sub

Private Function GuardJobStart() As Boolean
    ' Returns True if it's safe to start a new job. Otherwise warns the user.
    If IsJobRunning() Then
        MsgBox "Another job is already running. Click the STOP checkbox to halt it first.", _
               vbExclamation, "Stock Picker"
        GuardJobStart = False
    Else
        GuardJobStart = True
    End If
End Function

Sub RebuildInventory()
    If Not GuardJobStart() Then Exit Sub
    SetJobRunning True
    ClearStopRequest
    On Error GoTo cleanup
    RunPython "import stocks_report; stocks_report.button_rebuild_inventory()"
cleanup:
    SetJobRunning False
End Sub

Sub GetQuotes()
    If Not GuardJobStart() Then Exit Sub
    SetJobRunning True
    ClearStopRequest
    On Error GoTo cleanup
    RunPython "import stocks_report; stocks_report.button_get_quotes()"
cleanup:
    SetJobRunning False
End Sub
