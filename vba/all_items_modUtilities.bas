Option Explicit

' --- refresh ---------------------------------------------------------------
' Refreshes each connection synchronously. Returns only once every query has
' finished, so the calling Python code does not need a `time.sleep()` to wait
' for background queries.
'
' Performance toggles disable screen updates, automatic calculation, events,
' and alerts during the refresh; the original Application state is restored
' in the Cleanup block whether the Sub succeeded or raised an error.
'
' Side effect: each WorkbookConnection's BackgroundQuery flag is set to
' False and persists in the saved workbook. This is intentional.
Sub refresh()
    Dim conn As WorkbookConnection
    Dim prevCalc As Long

    On Error GoTo Cleanup

    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    For Each conn In ThisWorkbook.Connections
        On Error Resume Next
        conn.OLEDBConnection.BackgroundQuery = False
        On Error GoTo Cleanup
        conn.refresh
    Next conn

Cleanup:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = prevCalc
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "refresh error " & Err.Number & ": " & Err.Description, vbExclamation
    End If
End Sub


' --- sortAll ---------------------------------------------------------------
' Sorts the `Inactive_Listings` table on Sheet 1 by GS1 Type (descending) —
' the only sort op this workbook needs. Called by sibling automations after
' they refresh the Inactive Listings query so newest GS1 codes surface first.
'
' Performance toggles + Cleanup block as in the per-account workbook's sortAll.
Sub sortAll()
    Dim sh As Worksheet
    Dim tbl As ListObject
    Dim prevCalc As Long

    On Error GoTo Cleanup

    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Set sh = ThisWorkbook.Sheets(1)
    Set tbl = sh.ListObjects("Inactive_Listings")

    tbl.Sort.SortFields.Clear
    tbl.Sort.SortFields.Add2 _
        Key:=Range("Inactive_Listings[GS1 Type]"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With tbl.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

Cleanup:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = prevCalc
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "sortAll error " & Err.Number & ": " & Err.Description, vbExclamation
    End If
End Sub
