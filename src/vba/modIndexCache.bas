Attribute VB_Name = "modIndexCache"
Option Explicit
Option Private Module

' Exclusive session index backend:
'   Sheet  — onboard Database!tblFiles is non-empty (used for the whole session)
'   AccDB  — sheet empty and {workbook}\DB\LAN_Search_Index.accdb exists
'   Empty  — neither available
' No hybrid merge. Resolve once on warm / EnsureResolved.

Private Const SHEET_DATABASE As String = "Database"
Private Const TABLE_FILES As String = "tblFiles"
Private Const ACCDB_DIR As String = "DB"
Private Const ACCDB_FILE As String = "LAN_Search_Index.accdb"

Public Const BACKEND_EMPTY As String = "Empty"
Public Const BACKEND_SHEET As String = "Sheet"
Public Const BACKEND_ACCDB As String = "AccDB"

Private mBackend As String
Private mResolved As Boolean
Private mWorkbookPath As String

' Slim sheet cache (FilePath-based search only — no parent/descendant arrays)
Private mData As Variant
Private mRowCount As Long
Private mSheetLoaded As Boolean

Private mDriveByLabel As Object       ' Scripting.Dictionary: friendly label -> share UNC
Private mLabelByShare As Object       ' Scripting.Dictionary: share UNC -> friendly label

Private mLastGetWasHit As Boolean
Private mLastGetSeconds As Double
Private mLastLoadSeconds As Double

Public Function CurrentBackend() As String
    If Not mResolved Then
        CurrentBackend = BACKEND_EMPTY
    Else
        CurrentBackend = mBackend
    End If
End Function

Public Function IsIndexLoaded() As Boolean
    IsIndexLoaded = mResolved And (mBackend <> BACKEND_EMPTY)
End Function

Public Function IndexRowCount() As Long
    IndexRowCount = mRowCount
End Function

Public Function LastGetWasCacheHit() As Boolean
    LastGetWasCacheHit = mLastGetWasHit
End Function

Public Function LastGetSeconds() As Double
    LastGetSeconds = mLastGetSeconds
End Function

Public Function LastLoadSeconds() As Double
    LastLoadSeconds = mLastLoadSeconds
End Function

Public Function AccdbPath(Optional ByVal wb As Workbook = Nothing) As String
    Dim base As String
    If wb Is Nothing Then Set wb = ActiveWorkbook
    If wb Is Nothing Then
        AccdbPath = vbNullString
        Exit Function
    End If
    base = Trim$(wb.Path)
    If Len(base) = 0 Then
        AccdbPath = vbNullString
        Exit Function
    End If
    AccdbPath = base & "\" & ACCDB_DIR & "\" & ACCDB_FILE
End Function

Public Function AccdbExists(Optional ByVal wb As Workbook = Nothing) As Boolean
    Dim p As String
    p = AccdbPath(wb)
    If Len(p) = 0 Then
        AccdbExists = False
        Exit Function
    End If
    AccdbExists = (Len(Dir$(p)) > 0)
End Function

' Late-bound ACE connection to the AccDB beside the workbook.
Public Function OpenIndexConnection(Optional ByVal wb As Workbook = Nothing) As Object
    Dim cn As Object
    Dim p As String
    Dim errMsg As String

    p = AccdbPath(wb)
    If Len(p) = 0 Then
        Err.Raise vbObjectError + 100, "modIndexCache", _
                  "Workbook must be saved before opening AccDB (wb.Path is empty)."
    End If
    If Len(Dir$(p)) = 0 Then
        Err.Raise vbObjectError + 101, "modIndexCache", "AccDB not found: " & p
    End If

    Set cn = CreateObject("ADODB.Connection")
    On Error Resume Next
    cn.Open "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & p & ";"
    If Err.Number <> 0 Then
        errMsg = Err.Description
        Err.Clear
        cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & p & ";"
    End If
    If Err.Number <> 0 Then
        errMsg = Err.Description
        On Error GoTo 0
        Err.Raise vbObjectError + 102, "modIndexCache", _
                  "Could not open AccDB (install ACE matching Office bitness): " & errMsg
    End If
    On Error GoTo 0
    Set OpenIndexConnection = cn
End Function

Public Function SheetTblFilesNonEmpty(Optional ByVal wb As Workbook = Nothing) As Boolean
    Dim ws As Worksheet
    Dim tbl As ListObject

    SheetTblFilesNonEmpty = False
    If wb Is Nothing Then Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Function

    On Error Resume Next
    Set ws = wb.Worksheets(SHEET_DATABASE)
    If ws Is Nothing Then Exit Function
    Set tbl = ws.ListObjects(TABLE_FILES)
    On Error GoTo 0
    If tbl Is Nothing Then Exit Function
    If tbl.DataBodyRange Is Nothing Then Exit Function
    If tbl.ListRows.Count <= 0 Then Exit Function

    ' Treat a single blank FilePath row as empty
    If tbl.ListRows.Count = 1 Then
        If Len(Trim$(CStr(tbl.DataBodyRange.Cells(1, 1).Value2 & ""))) = 0 Then Exit Function
    End If
    SheetTblFilesNonEmpty = True
End Function

' Resolve exclusive backend once per session (or after Invalidate / SetIndexData).
Public Sub ResolveIndexBackend(Optional ByVal wb As Workbook = Nothing)
    Dim t0 As Double
    t0 = Timer
    mLastGetWasHit = mResolved

    If wb Is Nothing Then Set wb = ActiveWorkbook
    If wb Is Nothing Then
        InvalidateIndex
        mLastLoadSeconds = ElapsedSeconds(t0)
        Exit Sub
    End If

    If mResolved And Len(mWorkbookPath) > 0 Then
        If StrComp(mWorkbookPath, wb.FullName, vbTextCompare) = 0 Then
            mLastGetSeconds = ElapsedSeconds(t0)
            RefreshDashboardDriveList wb
            Exit Sub
        End If
    End If

    InvalidateIndex
    mWorkbookPath = wb.FullName

    If SheetTblFilesNonEmpty(wb) Then
        mBackend = BACKEND_SHEET
        LoadSheetSlim wb
        BuildDriveMapsFromSheet
        mResolved = True
        mLastLoadSeconds = ElapsedSeconds(t0)
        Debug.Print "INDEX: ResolveIndexBackend=Sheet rows=" & CStr(mRowCount) & _
                    " " & FormatSeconds(mLastLoadSeconds)
        RefreshDashboardDriveList wb
        Exit Sub
    End If

    If AccdbExists(wb) Then
        mBackend = BACKEND_ACCDB
        mRowCount = AccdbRowCount(wb)
        mSheetLoaded = False
        mData = Empty
        BuildDriveMapsFromAccdb wb
        mResolved = True
        mLastLoadSeconds = ElapsedSeconds(t0)
        Debug.Print "INDEX: ResolveIndexBackend=AccDB rows=" & CStr(mRowCount) & _
                    " " & FormatSeconds(mLastLoadSeconds)
        RefreshDashboardDriveList wb
        Exit Sub
    End If

    mBackend = BACKEND_EMPTY
    mResolved = True
    mRowCount = 0
    mLastLoadSeconds = ElapsedSeconds(t0)
    Debug.Print "INDEX: ResolveIndexBackend=Empty " & FormatSeconds(mLastLoadSeconds)
    RefreshDashboardDriveList wb
End Sub

Public Sub EnsureLoaded(Optional ByVal wb As Workbook = Nothing)
    Dim t0 As Double
    t0 = Timer
    mLastGetWasHit = mResolved And (mBackend <> BACKEND_EMPTY)
    If Not mResolved Then
        ResolveIndexBackend wb
    Else
        RefreshDashboardDriveList wb
    End If
    mLastGetSeconds = ElapsedSeconds(t0)
End Sub

Public Sub WarmIndexIfNeeded(Optional ByVal wb As Workbook = Nothing)
    If mResolved Then
        Debug.Print "INDEX: already resolved backend=" & mBackend & " rows=" & CStr(mRowCount)
        Exit Sub
    End If
    Debug.Print "INDEX: WarmIndexIfNeeded resolving backend..."
    ResolveIndexBackend wb
End Sub

' Seed sheet backend after VBA ingest write.
Public Sub SetIndexData(ByRef source As Variant, ByVal rowCount As Long)
    Dim t0 As Double
    If rowCount <= 0 Or Not IsArray(source) Then
        InvalidateIndex
        Exit Sub
    End If
    mData = source
    mRowCount = rowCount
    mSheetLoaded = True
    mBackend = BACKEND_SHEET
    mResolved = True
    mLastLoadSeconds = 0
    t0 = Timer
    BuildDriveMapsFromSheet
    Debug.Print "INDEX: SetIndexData Sheet rows=" & CStr(mRowCount) & _
                " drives=" & FormatSeconds(ElapsedSeconds(t0))
    RefreshDashboardDriveList ActiveWorkbook
End Sub

Public Sub InvalidateIndex()
    On Error Resume Next
    Erase mData
    Set mDriveByLabel = Nothing
    Set mLabelByShare = Nothing
    On Error GoTo 0
    mData = Empty
    mRowCount = 0
    mSheetLoaded = False
    mResolved = False
    mBackend = BACKEND_EMPTY
    mWorkbookPath = vbNullString
    Debug.Print "INDEX: invalidated"
End Sub

' Slim path/LIKE-equivalent scan of onboard sheet cache.
Public Sub CollectPathHits(ByVal term1 As String, ByVal op As String, ByVal term2 As String, _
                           ByVal flexible As Boolean, _
                           ByVal includeFiles As Boolean, ByVal includeFolders As Boolean, _
                           ByVal hasSizeFilter As Boolean, ByVal sizeOp As String, ByVal sizeMb As Double, _
                           ByVal hits As Object, _
                           Optional ByVal shareFilter As String = "")
    Dim i As Long
    Dim path As String
    Dim et As String
    Dim isFolder As Boolean
    Dim sizeVal As Double
    Dim dt As Variant
    Dim term1N As String
    Dim term2N As String
    Dim hay As String
    Dim hasTerm2 As Boolean
    Dim has1 As Boolean
    Dim has2 As Boolean
    Dim matched As Boolean
    Dim filterShare As String
    Dim share As String

    If mBackend <> BACKEND_SHEET Or Not mSheetLoaded Or mRowCount = 0 Then Exit Sub

    term1 = Trim$(term1)
    term2 = Trim$(term2)
    op = UCase$(Trim$(op))
    hasTerm2 = (Len(term2) > 0 And Len(op) > 0)
    filterShare = Trim$(shareFilter)

    If flexible Then
        term1N = modPathUtil.NormalizeForMatch(term1)
        If hasTerm2 Then term2N = modPathUtil.NormalizeForMatch(term2)
        If Len(term1N) = 0 Then Exit Sub
    End If

    For i = 1 To mRowCount
        path = CStr(mData(i, 1) & "")
        If Right$(path, 1) = "\" Then path = Left$(path, Len(path) - 1)
        If Len(path) = 0 Then GoTo NextRow

        If Len(filterShare) > 0 Then
            share = modPathUtil.UncShareRoot(path)
            If StrComp(share, filterShare, vbTextCompare) <> 0 Then GoTo NextRow
        End If

        et = UCase$(Trim$(CStr(mData(i, 4) & "")))
        If Len(et) = 0 Then et = "FILE"
        isFolder = (et = "FOLDER")
        If isFolder Then
            If Not includeFolders Then GoTo NextRow
        Else
            If Not includeFiles Then GoTo NextRow
        End If

        If flexible Then
            hay = modPathUtil.NormalizeForMatch(path)
            has1 = (InStr(1, hay, term1N, vbBinaryCompare) > 0)
            If hasTerm2 Then
                has2 = (InStr(1, hay, term2N, vbBinaryCompare) > 0)
                matched = CombineMatch(has1, has2, op)
            Else
                matched = has1
            End If
        Else
            hay = path
            has1 = (InStr(1, hay, term1, vbTextCompare) > 0)
            If hasTerm2 Then
                has2 = (InStr(1, hay, term2, vbTextCompare) > 0)
                matched = CombineMatch(has1, has2, op)
            Else
                matched = has1
            End If
        End If
        If Not matched Then GoTo NextRow

        sizeVal = modPathUtil.CoerceStoredSizeMb(mData(i, 3))
        If hasSizeFilter Then
            If Not SizeMatchesLocal(sizeVal, sizeOp, sizeMb) Then GoTo NextRow
        End If

        dt = mData(i, 2)
        If Not hits.Exists(path) Then
            hits.Add path, Array(dt, sizeVal, isFolder, 0&, 0&)
        End If
NextRow:
    Next i
End Sub

Public Function ResolveDriveShareFilter(ByVal label As String) As String
    Dim v As String
    v = Trim$(label)
    If Len(v) = 0 Or StrComp(v, "All", vbTextCompare) = 0 Then
        ResolveDriveShareFilter = vbNullString
        Exit Function
    End If
    If mDriveByLabel Is Nothing Then
        ResolveDriveShareFilter = vbNullString
        Exit Function
    End If
    If mDriveByLabel.Exists(v) Then
        ResolveDriveShareFilter = CStr(mDriveByLabel(v))
    Else
        ResolveDriveShareFilter = modPathUtil.UncShareRoot(v)
    End If
End Function

Public Function FriendlyLabelForShare(ByVal shareOrPath As String) As String
    Dim share As String
    Dim label As String

    share = Trim$(shareOrPath)
    If Len(share) = 0 Then
        FriendlyLabelForShare = vbNullString
        Exit Function
    End If
    If Left$(share, 2) = "\\" Or (Len(share) >= 2 And Mid$(share, 2, 1) = ":") Then
        share = modPathUtil.UncShareRoot(share)
    End If
    If Len(share) = 0 Then
        FriendlyLabelForShare = vbNullString
        Exit Function
    End If

    If Not mLabelByShare Is Nothing Then
        If mLabelByShare.Exists(share) Then
            FriendlyLabelForShare = CStr(mLabelByShare(share))
            Exit Function
        End If
    End If

    label = modPathUtil.FriendlyDriveLabel(share)
    If Len(label) = 0 Then label = share
    FriendlyLabelForShare = label
End Function

Public Sub RefreshDashboardDriveList(Optional ByVal wb As Workbook = Nothing)
    Dim map As clsSheetMap
    Dim ws As Worksheet
    Dim cell As Range
    Dim k As Variant
    Dim i As Long
    Dim n As Long
    Dim labels() As String
    Dim prev As String
    Dim keep As String
    Dim formula As String
    Dim activeWas As Worksheet

    On Error GoTo FailQuiet
    If wb Is Nothing Then Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub

    Set map = New clsSheetMap
    map.Init wb
    map.UnprotectDashboard
    Set ws = map.Dashboard
    Set cell = ws.Range(map.DriveFilterAddress)
    On Error Resume Next
    If cell.MergeCells Then Set cell = cell.MergeArea.Cells(1, 1)
    On Error GoTo FailQuiet
    prev = Trim$(CStr(cell.Value & ""))

    n = 0
    If Not mDriveByLabel Is Nothing Then n = mDriveByLabel.Count
    ReDim labels(1 To n + 1)
    labels(1) = "All"
    i = 1
    If n > 0 Then
        For Each k In mDriveByLabel.Keys
            i = i + 1
            labels(i) = CStr(k)
        Next k
        If n > 1 Then SortStringsAz labels, 2, n + 1
    End If

    formula = labels(1)
    For i = 2 To n + 1
        If Len(labels(i)) > 0 Then formula = formula & "," & labels(i)
    Next i

    On Error Resume Next
    Set activeWas = ActiveSheet
    ws.Activate
    cell.Validation.Delete
    Err.Clear
    cell.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
                        Formula1:=formula
    On Error GoTo FailQuiet

    With cell.Validation
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = False
        .ShowError = True
        .ErrorTitle = "Drive"
        .ErrorMessage = "Pick a drive from the list (or All)."
    End With

    keep = "All"
    If Len(prev) > 0 And StrComp(prev, "All", vbTextCompare) <> 0 Then
        If Not mDriveByLabel Is Nothing Then
            If mDriveByLabel.Exists(prev) Then keep = prev
        End If
    End If
    cell.Value = keep

    If Not activeWas Is Nothing Then
        On Error Resume Next
        If StrComp(activeWas.Name, ws.Name, vbTextCompare) = 0 Or activeWas.Visible <> xlSheetVisible Then
            ws.Activate
        Else
            activeWas.Activate
        End If
        On Error GoTo FailQuiet
    End If

    Debug.Print "INDEX: drive dropdown items=" & CStr(n + 1) & " backend=" & mBackend
    Exit Sub
FailQuiet:
    Debug.Print "INDEX: RefreshDashboardDriveList skipped Err=" & CStr(Err.Number) & " " & Err.Description
End Sub

' --- Private helpers ---

Private Sub LoadSheetSlim(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim data As Variant
    Dim n As Long
    Dim cols As Long
    Dim r As Long
    Dim normalized() As Variant

    mSheetLoaded = False
    mData = Empty
    mRowCount = 0

    On Error Resume Next
    Set ws = wb.Worksheets(SHEET_DATABASE)
    Set tbl = ws.ListObjects(TABLE_FILES)
    On Error GoTo 0
    If tbl Is Nothing Then Exit Sub
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    If tbl.ListRows.Count = 0 Then Exit Sub

    data = tbl.DataBodyRange.Value
    If Not IsArray(data) Then Exit Sub

    n = UBound(data, 1)
    cols = UBound(data, 2)
    ReDim normalized(1 To n, 1 To 4)
    For r = 1 To n
        normalized(r, 1) = data(r, 1)
        If cols >= 2 Then normalized(r, 2) = data(r, 2) Else normalized(r, 2) = Empty
        If cols >= 3 Then normalized(r, 3) = data(r, 3) Else normalized(r, 3) = 0.01
        If cols >= 4 And Len(Trim$(CStr(data(r, 4) & ""))) > 0 Then
            normalized(r, 4) = UCase$(Trim$(CStr(data(r, 4))))
        Else
            normalized(r, 4) = "FILE"
        End If
    Next r

    mData = normalized
    mRowCount = n
    mSheetLoaded = True
End Sub

Private Function AccdbRowCount(ByVal wb As Workbook) As Long
    Dim cn As Object
    Dim rs As Object
    AccdbRowCount = 0
    On Error GoTo Fail
    Set cn = OpenIndexConnection(wb)
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT COUNT(*) AS Cnt FROM tblFiles", cn
    If Not rs.EOF Then AccdbRowCount = CLng(rs.Fields(0).Value & 0)
    rs.Close
    cn.Close
    Exit Function
Fail:
    AccdbRowCount = 0
End Function

Private Sub BuildDriveMapsFromSheet()
    Dim i As Long
    Dim path As String
    InitDriveMaps
    If Not mSheetLoaded Or mRowCount = 0 Then Exit Sub
    For i = 1 To mRowCount
        path = CStr(mData(i, 1) & "")
        AddShareFromPath path
    Next i
End Sub

Private Sub BuildDriveMapsFromAccdb(ByVal wb As Workbook)
    Dim cn As Object
    Dim rs As Object
    Dim share As String
    Dim sql As String

    InitDriveMaps
    On Error GoTo Fail
    Set cn = OpenIndexConnection(wb)
    Set rs = CreateObject("ADODB.Recordset")
    ' Distinct \\server\share from UNC FilePath (3rd backslash ends the share)
    sql = "SELECT DISTINCT Left([FilePath], InStr(InStr(3,[FilePath],'\')+1,[FilePath],'\')-1) AS ShareRoot " & _
          "FROM tblFiles WHERE Left([FilePath],2)='\\' AND InStr(3,[FilePath],'\')>0 " & _
          "AND InStr(InStr(3,[FilePath],'\')+1,[FilePath],'\')>0"
    rs.Open sql, cn
    Do While Not rs.EOF
        share = Trim$(CStr(rs.Fields(0).Value & ""))
        If Len(share) > 0 Then AddShareFromPath share & "\"
        rs.MoveNext
    Loop
    rs.Close
    cn.Close
    Exit Sub
Fail:
    Debug.Print "INDEX: BuildDriveMapsFromAccdb failed " & Err.Description
End Sub

Private Sub InitDriveMaps()
    Set mDriveByLabel = CreateObject("Scripting.Dictionary")
    mDriveByLabel.CompareMode = 1
    Set mLabelByShare = CreateObject("Scripting.Dictionary")
    mLabelByShare.CompareMode = 1
End Sub

Private Sub AddShareFromPath(ByVal path As String)
    Dim share As String
    Dim label As String
    Dim unique As String
    Dim suffix As Long

    path = Trim$(path)
    If Len(path) = 0 Then Exit Sub
    share = modPathUtil.UncShareRoot(path)
    If Len(share) = 0 Then Exit Sub
    If mLabelByShare.Exists(share) Then Exit Sub

    label = modPathUtil.FriendlyDriveLabel(share)
    If Len(label) = 0 Then label = share
    unique = label
    suffix = 2
    Do While mDriveByLabel.Exists(unique)
        unique = label & " (" & CStr(suffix) & ")"
        suffix = suffix + 1
    Loop
    mDriveByLabel.Add unique, share
    mLabelByShare.Add share, unique
End Sub

Private Sub SortStringsAz(ByRef values() As String, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long
    Dim j As Long
    Dim pivot As String
    Dim tmp As String
    i = lo
    j = hi
    pivot = values((lo + hi) \ 2)
    Do While i <= j
        Do While StrComp(values(i), pivot, vbTextCompare) < 0
            i = i + 1
        Loop
        Do While StrComp(values(j), pivot, vbTextCompare) > 0
            j = j - 1
        Loop
        If i <= j Then
            tmp = values(i)
            values(i) = values(j)
            values(j) = tmp
            i = i + 1
            j = j - 1
        End If
    Loop
    If lo < j Then SortStringsAz values, lo, j
    If i < hi Then SortStringsAz values, i, hi
End Sub

Private Function CombineMatch(ByVal has1 As Boolean, ByVal has2 As Boolean, ByVal op As String) As Boolean
    Select Case op
        Case "AND"
            CombineMatch = has1 And has2
        Case "OR"
            CombineMatch = has1 Or has2
        Case "NOT"
            CombineMatch = has1 And (Not has2)
        Case Else
            CombineMatch = has1
    End Select
End Function

Private Function SizeMatchesLocal(ByVal fileMb As Double, ByVal sizeOp As String, ByVal sizeMb As Double) As Boolean
    Select Case sizeOp
        Case ">"
            SizeMatchesLocal = (fileMb > sizeMb)
        Case "<"
            SizeMatchesLocal = (fileMb < sizeMb)
        Case Else
            SizeMatchesLocal = True
    End Select
End Function

Private Function ElapsedSeconds(ByVal startTimer As Double) As Double
    Dim nowT As Double
    nowT = Timer
    If nowT >= startTimer Then
        ElapsedSeconds = nowT - startTimer
    Else
        ElapsedSeconds = (86400# - startTimer) + nowT
    End If
End Function

Private Function FormatSeconds(ByVal seconds As Double) As String
    FormatSeconds = Format$(seconds, "0.000") & "s"
End Function
