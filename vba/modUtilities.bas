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
    'Sort each per-account table by Date (desc) then Order ID (asc).
    'First 6 sheets are the account tables; their ListObject name is the
    'sheet name with spaces replaced by underscores.
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


Sub sortStatus()
    'Sort each per-account table by Status (desc) only — used to surface
    'open feedback items at the top of each sheet.
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
