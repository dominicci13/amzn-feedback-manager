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
' Sorts each per-account table by Date (descending) then Order ID (ascending).
' The first 6 sheets are the per-account tables; their ListObject name is
' derived from the sheet name with spaces replaced by underscores (e.g.
' "SellerOrg" → "Account_A"). Each sheet's selection is parked on B4
' after sorting so the table opens at the top-left next time.
'
' Performance toggles match the canonical Pending-Offers pattern; the
' Cleanup block restores Application state whether the Sub succeeded or
' raised an error.
Sub sortAll()
    Dim sh As Worksheet
    Dim tbl As ListObject
    Dim tableName As String
    Dim i As Long
    Dim prevCalc As Long

    On Error GoTo Cleanup

    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    For i = 1 To 6
        Set sh = ThisWorkbook.sheets(i)
        tableName = Replace(sh.Name, " ", "_")
        Set tbl = sh.ListObjects(tableName)

        tbl.Sort.SortFields.Clear
        tbl.Sort.SortFields.Add2 _
            Key:=Range(tableName & "[Date]"), _
            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        tbl.Sort.SortFields.Add2 _
            Key:=Range(tableName & "[Order ID]"), _
            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With tbl.Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With

        sh.Activate
        sh.Range("B4").Select
    Next i
    ThisWorkbook.sheets(1).Activate

Cleanup:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = prevCalc
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "sortAll error " & Err.Number & ": " & Err.Description, vbExclamation
    End If
End Sub


' --- sortStatus ------------------------------------------------------------
' Re-sorts each per-account table by Status (descending only) to surface
' open feedback items at the top of every sheet. Same per-sheet ListObject
' lookup as sortAll. Called after the Python step writes new Status values
' so the next sortAll preserves the "open at top" visual.
'
' Performance toggles + Cleanup block as in sortAll.
Sub sortStatus()
    Dim sh As Worksheet
    Dim tbl As ListObject
    Dim tableName As String
    Dim i As Long
    Dim prevCalc As Long

    On Error GoTo Cleanup

    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    For i = 1 To 6
        Set sh = ThisWorkbook.sheets(i)
        tableName = Replace(sh.Name, " ", "_")
        Set tbl = sh.ListObjects(tableName)

        tbl.Sort.SortFields.Clear
        tbl.Sort.SortFields.Add2 _
            Key:=Range(tableName & "[Status]"), _
            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        With tbl.Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With

        sh.Activate
        sh.Range("B4").Select
    Next i
    ThisWorkbook.sheets(1).Activate

Cleanup:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = prevCalc
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "sortStatus error " & Err.Number & ": " & Err.Description, vbExclamation
    End If
End Sub
