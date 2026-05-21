Option Explicit


Sub refresh()
    Dim conn As WorkbookConnection
    Dim prevCalc As Long

    On Error GoTo Cleanup

    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    'Force synchronous refresh on every Power Query connection.
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


Sub sortAll()
    'Sort Inactive_Listings by GS1 Type desc — the only sort op this workbook needs.
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
