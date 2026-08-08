Attribute VB_Name = "modEntryPoints"
Option Explicit

' Button macros:
'   RunIndexFromIngestion
'   RunSearchFromDashboard
' WarmIndexCache is Public for Application.OnTime (Workbook_Open) — not for buttons.
'
' Helpers: modPathUtil, modIndexCache (Option Private Module).

Public Sub RunIndexFromIngestion()
    Dim map As clsSheetMap
    Dim appState As clsExcelAppState
    Dim indexer As clsFileIndexer

    Set map = ReadySheetMap()
    Set appState = New clsExcelAppState
    Set indexer = New clsFileIndexer
    indexer.Init map, appState
    indexer.RunFromIngestion
End Sub

Public Sub RunSearchFromDashboard()
    Dim map As clsSheetMap
    Dim appState As clsExcelAppState
    Dim searcher As clsFileSearch

    Set map = ReadySheetMap()
    Set appState = New clsExcelAppState
    Set searcher = New clsFileSearch
    searcher.Init map, appState
    searcher.RunFromDashboard
End Sub

' Deferred from ThisWorkbook.Workbook_Open via Application.OnTime.
Public Sub WarmIndexCache()
    Dim map As clsSheetMap
    On Error GoTo QuietFail
    Debug.Print "INDEX CACHE: WarmIndexCache OnTime fired"
    Application.EnableEvents = True
    modIndexCache.WarmIndexIfNeeded ActiveWorkbook
    Set map = ReadySheetMap()
    map.EnsureResultsDisplayGrid
    map.ProtectDashboardResults
    ResetStatusBar
    Exit Sub
QuietFail:
    Debug.Print "INDEX CACHE: WarmIndexCache failed Err=" & CStr(Err.Number) & " " & Err.Description
    ResetStatusBar
End Sub

' Double-click File name → open containing folder (Explorer).
' Cancels edit / protect-warning for any locked Dashboard cell.
Public Function HandleDashboardResultDoubleClick(ByVal Sh As Object, ByVal Target As Range, ByRef Cancel As Boolean) As Boolean
    Dim map As clsSheetMap
    Dim r As Long
    Dim topCol As Long
    Dim unc As String
    Dim firstData As Long
    Dim lastData As Long
    Dim cellLocked As Boolean

    HandleDashboardResultDoubleClick = False
    On Error GoTo Fail

    If Sh Is Nothing Or Target Is Nothing Then Exit Function
    If StrComp(Sh.Name, "Dashboard", vbTextCompare) <> 0 Then Exit Function

    ' Never enter edit mode / show protect dialog on locked cells
    On Error Resume Next
    cellLocked = Sh.Cells(Target.Row, Target.Column).Locked
    On Error GoTo Fail
    If cellLocked Then
        Cancel = True
        HandleDashboardResultDoubleClick = True
    End If

    Set map = New clsSheetMap
    map.Init Sh.Parent

    firstData = map.ResultsFirstDataRow
    lastData = firstData + map.ResultsMaxRows - 1
    r = Target.Row
    If r < firstData Or r > lastData Then Exit Function

    Cancel = True
    HandleDashboardResultDoubleClick = True

    ' Merged File band A:K always reports Column = 1 — do NOT touch MergeArea (1004 when protected)
    topCol = Target.Column

    If topCol < map.ResultsFileCol Or topCol > map.ResultsFileColEnd Then
        Debug.Print "SEARCH: double-click row=" & CStr(r) & " col=" & CStr(topCol) & " (not File; edit blocked)"
        Exit Function
    End If

    unc = Trim$(Replace(CStr(Sh.Cells(r, map.ResultsLinkCol).Value2 & ""), "/", "\"))
    If Len(unc) = 0 Then
        Debug.Print "SEARCH: double-click File row=" & CStr(r) & " but UNC link col empty"
        Exit Function
    End If

    Debug.Print "SEARCH: open folder " & unc
    On Error Resume Next
    Shell "explorer.exe """ & unc & """", vbNormalFocus
    If Err.Number <> 0 Then
        Debug.Print "SEARCH: Shell failed Err=" & CStr(Err.Number) & " " & Err.Description
        Err.Clear
        Sh.Parent.FollowHyperlink Address:=unc
        If Err.Number <> 0 Then
            Debug.Print "SEARCH: FollowHyperlink failed Err=" & CStr(Err.Number) & " " & Err.Description
        End If
    End If
    On Error GoTo 0
    Exit Function

Fail:
    Debug.Print "SEARCH: double-click handler Err=" & CStr(Err.Number) & " " & Err.Description
    Cancel = True
End Function

Public Sub ResetStatusBar()
    On Error Resume Next
    Application.DisplayStatusBar = True
    Application.StatusBar = vbNullString
    Application.StatusBar = False
    On Error GoTo 0
End Sub

Private Function ReadySheetMap() As clsSheetMap
    Dim map As clsSheetMap
    Set map = New clsSheetMap
    map.Init ActiveWorkbook
    ' Do NOT call EnsureLayoutLabels here — it was overwriting the user's Ingestion UI.
    map.EnsureModernSheetNames
    Set ReadySheetMap = map
End Function

Private Sub RunRefreshIngestedList()
    Dim map As clsSheetMap
    Dim appState As clsExcelAppState
    Dim indexer As clsFileIndexer

    Set map = ReadySheetMap()
    Set appState = New clsExcelAppState
    Set indexer = New clsFileIndexer
    indexer.Init map, appState
    indexer.RefreshIngestedListFromDatabase
End Sub
