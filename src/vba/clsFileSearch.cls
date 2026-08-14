VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFileSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Searches the exclusive session backend (Sheet or AccDB):
' - Sheet: slim FilePath scan via modIndexCache.CollectPathHits
' - AccDB: late-bound ADODB SQL with FilePath LIKE / size / EntryType filters

Private mMap As clsSheetMap
Private mApp As clsExcelAppState
Private mLastHitCount As Long
Private mLastDisplayedCount As Long

' hit meta: Array(date, sizeMb, isFolder, fileCount, folderCount)

Public Sub Init(ByVal map As clsSheetMap, ByVal appState As clsExcelAppState)
    Set mMap = map
    Set mApp = appState
End Sub

Public Sub RunFromDashboard()
    Dim hits As Object
    Dim calcPrev As XlCalculation
    Dim anyCriteria As Boolean
    Dim rowCount As Long
    Dim tAll As Double
    Dim tPhase As Double
    Dim secGet As Double
    Dim secScan As Double
    Dim secWrite As Double
    Dim cacheHit As Boolean
    Dim backend As String

    tAll = Timer
    calcPrev = Application.Calculation
    On Error GoTo Fail
    mApp.BeginFast
    Application.Calculation = xlCalculationManual

    Set hits = CreateObject("Scripting.Dictionary")
    hits.CompareMode = 1 ' TextCompare

    Debug.Print "========== SEARCH TIMING START =========="
    Debug.Print "SEARCH: resolved before get? " & CStr(modIndexCache.IsIndexLoaded) & _
                " backend=" & modIndexCache.CurrentBackend & _
                " rows=" & CStr(modIndexCache.IndexRowCount)

    tPhase = Timer
    modIndexCache.EnsureLoaded ActiveWorkbook
    secGet = ElapsedSeconds(tPhase)
    cacheHit = modIndexCache.LastGetWasCacheHit
    rowCount = modIndexCache.IndexRowCount
    backend = modIndexCache.CurrentBackend

    Debug.Print "SEARCH: EnsureLoaded backend=" & backend & " rows=" & CStr(rowCount) & _
                " hit=" & CStr(cacheHit) & " elapsed=" & FormatSeconds(secGet)

    If rowCount = 0 Or backend = modIndexCache.BACKEND_EMPTY Then
        ClearResults
        WriteSearchStats 0, 0, 0#
        mMap.ProtectDashboardResults
        Application.Calculation = calcPrev
        mApp.Restore
        mMap.FocusDashboardResultsHome
        Debug.Print "SEARCH: aborted empty index total=" & FormatSeconds(ElapsedSeconds(tAll))
        MsgBox "No index available." & vbCrLf & vbCrLf & _
               "Onboard Database!tblFiles is empty and no AccDB was found at:" & vbCrLf & _
               modIndexCache.AccdbPath(ActiveWorkbook) & vbCrLf & vbCrLf & _
               "Crawl into the workbook (small) or AccDB (large), then reopen.", _
               vbExclamation, "LAN Search - Dashboard"
        Exit Sub
    End If

    tPhase = Timer
    SearchIndex hits, anyCriteria
    secScan = ElapsedSeconds(tPhase)
    Debug.Print "SEARCH: SearchIndex total=" & FormatSeconds(secScan) & _
                " hits=" & CStr(hits.Count) & " anyCriteria=" & CStr(anyCriteria)

    If Not anyCriteria Then
        ClearResults
        WriteSearchStats 0, 0, 0#
        mMap.ProtectDashboardResults
        Application.Calculation = calcPrev
        mApp.Restore
        mMap.FocusDashboardResultsHome
        Debug.Print "SEARCH: aborted no criteria total=" & FormatSeconds(ElapsedSeconds(tAll))
        MsgBox "Enter at least one Key Word in the criteria table.", vbExclamation, "LAN Search - Dashboard"
        Exit Sub
    End If

    tPhase = Timer
    WriteResults hits
    secWrite = ElapsedSeconds(tPhase)
    Debug.Print "SEARCH: WriteResults total=" & FormatSeconds(secWrite)

    Application.Calculation = calcPrev
    mApp.Restore
    mMap.FocusDashboardResultsHome

    Debug.Print "SEARCH SUMMARY: get=" & FormatSeconds(secGet) & _
                " scan=" & FormatSeconds(secScan) & _
                " write=" & FormatSeconds(secWrite) & _
                " TOTAL=" & FormatSeconds(ElapsedSeconds(tAll)) & _
                " backend=" & backend & " rows=" & CStr(rowCount) & " hits=" & CStr(hits.Count)
    Debug.Print "========== SEARCH TIMING END =========="

    If mLastDisplayedCount > 0 And mLastDisplayedCount < mLastHitCount Then
        MsgBox "Found " & Format$(mLastHitCount, "#,##0") & " result(s)." & vbCrLf & vbCrLf & _
               "Showing the first " & Format$(mLastDisplayedCount, "#,##0") & _
               ". Narrow keywords or pick a drive for fewer hits." & vbCrLf & vbCrLf & _
               "Double-click a File name to open its folder.", _
               vbInformation, "LAN Search - Dashboard"
    Else
        MsgBox "Found " & Format$(hits.Count, "#,##0") & " result(s)." & vbCrLf & vbCrLf & _
               "Double-click a File name to open its folder.", _
               vbInformation, "LAN Search - Dashboard"
    End If
    mMap.FocusDashboardResultsHome
    Exit Sub

Fail:
    Dim errDesc As String
    errDesc = Err.Description
    On Error Resume Next
    Application.Calculation = calcPrev
    mApp.Restore
    mMap.FocusDashboardResultsHome
    Debug.Print "SEARCH FAILED: " & errDesc & " after " & FormatSeconds(ElapsedSeconds(tAll))
    On Error GoTo 0
    MsgBox "Search failed: " & errDesc, vbCritical, "LAN Search - Dashboard"
End Sub

Private Sub SearchIndex(ByVal hits As Object, ByRef anyCriteria As Boolean)
    Dim fileTerm As String
    Dim term1 As String
    Dim term2 As String
    Dim op As String
    Dim sizeOp As String
    Dim sizeMb As Double
    Dim hasSizeFilter As Boolean
    Dim flexible As Boolean
    Dim foldersMode As String
    Dim includeFiles As Boolean
    Dim includeFolders As Boolean
    Dim shareFilter As String
    Dim backend As String
    Dim cn As Object
    Dim tScan As Double
    Dim fileHitsBefore As Long
    Dim extraAnd As String

    flexible = mMap.FlexibleMatch
    foldersMode = mMap.SearchFoldersMode
    includeFiles = (foldersMode <> "Only")
    includeFolders = (foldersMode = "Yes" Or foldersMode = "Only")
    shareFilter = modIndexCache.ResolveDriveShareFilter(mMap.DriveFilterLabel)
    backend = modIndexCache.CurrentBackend

    fileTerm = mMap.FileSearchTerm
    term1 = mMap.CriteriaTerm1
    op = mMap.CriteriaOp
    term2 = mMap.CriteriaTerm2
    sizeOp = mMap.CriteriaSizeOp
    hasSizeFilter = ParseSizeFilter(sizeOp, mMap.CriteriaSizeMbValue, sizeMb)
    extraAnd = vbNullString

    ' File Search (A5) is the main box. Contains row is optional advanced filter.
    If Len(fileTerm) > 0 And Len(term1) > 0 Then
        extraAnd = fileTerm
        ' term1/op/term2 stay as criteria Contains clause
    ElseIf Len(fileTerm) > 0 Then
        term1 = fileTerm
        op = vbNullString
        term2 = vbNullString
    End If

    Debug.Print "SEARCH: backend=" & backend & " foldersMode=" & foldersMode & _
                " flexible=" & CStr(flexible) & " drive=" & mMap.DriveFilterLabel & _
                " shareFilter=" & shareFilter & " n=" & CStr(modIndexCache.IndexRowCount) & _
                " term1=" & term1 & " extraAnd=" & extraAnd

    If Len(term1) = 0 Then
        anyCriteria = False
        Exit Sub
    End If
    If op = "NOT" And Len(term2) = 0 Then
        Err.Raise vbObjectError + 1, , "NOT requires a second Contains keyword."
    End If
    If Len(sizeOp) > 0 And Not hasSizeFilter Then
        Err.Raise vbObjectError + 2, , "Size filter incomplete (pick Over or Under and enter MB)."
    End If

    anyCriteria = True

    If backend = modIndexCache.BACKEND_ACCDB Then
        Set cn = modIndexCache.OpenIndexConnection(ActiveWorkbook)
    End If

    fileHitsBefore = hits.Count
    tScan = Timer

    If backend = modIndexCache.BACKEND_ACCDB Then
        SearchAccdbCriteria cn, hits, term1, op, term2, flexible, _
            includeFiles, includeFolders, hasSizeFilter, sizeOp, sizeMb, shareFilter, extraAnd
    Else
        modIndexCache.CollectPathHits term1, op, term2, flexible, _
            includeFiles, includeFolders, hasSizeFilter, sizeOp, sizeMb, _
            hits, shareFilter, extraAnd
    End If

    Debug.Print "SEARCH: scan=" & FormatSeconds(ElapsedSeconds(tScan)) & _
                " hits+" & CStr(hits.Count - fileHitsBefore)

    If Not cn Is Nothing Then
        On Error Resume Next
        cn.Close
        On Error GoTo 0
    End If
End Sub

Private Sub SearchAccdbCriteria(ByVal cn As Object, ByVal hits As Object, _
                                ByVal term1 As String, ByVal op As String, ByVal term2 As String, _
                                ByVal flexible As Boolean, _
                                ByVal includeFiles As Boolean, ByVal includeFolders As Boolean, _
                                ByVal hasSizeFilter As Boolean, ByVal sizeOp As String, ByVal sizeMb As Double, _
                                ByVal shareFilter As String, _
                                Optional ByVal extraAndTerm As String = "")
    Dim rs As Object
    Dim sql As String
    Dim whereSql As String
    Dim path As String
    Dim et As String
    Dim isFolder As Boolean
    Dim sizeVal As Double
    Dim dt As Variant
    Dim t1 As String
    Dim t2 As String
    Dim extraN As String

    t1 = term1
    t2 = term2
    extraN = Trim$(extraAndTerm)
    If flexible Then
        t1 = modPathUtil.NormalizeForMatch(term1)
        If Len(Trim$(term2)) > 0 Then t2 = modPathUtil.NormalizeForMatch(term2)
        If Len(extraN) > 0 Then extraN = modPathUtil.NormalizeForMatch(extraN)
        If Len(t1) = 0 Then Exit Sub
    End If

    whereSql = BuildAccdbWhere(t1, op, t2, includeFiles, includeFolders, _
                               hasSizeFilter, sizeOp, sizeMb, shareFilter, extraN)
    If Len(Trim$(whereSql)) = 0 Or whereSql = "()" Then
        Err.Raise vbObjectError + 3, , "AccDB search produced an empty WHERE clause."
    End If
    sql = "SELECT [FilePath], [FileDate], [SizeMB], [EntryType] FROM [tblFiles] WHERE " & whereSql
    Debug.Print "SEARCH: AccDB SQL=" & Left$(sql, 500)

    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3 ' adUseClient
    rs.Open sql, cn

    Do While Not rs.EOF
        path = Trim$(CStr(rs.Fields("FilePath").Value & ""))
        If Right$(path, 1) = "\" Then path = Left$(path, Len(path) - 1)
        If Len(path) > 0 Then
            et = UCase$(Trim$(CStr(rs.Fields("EntryType").Value & "")))
            If Len(et) = 0 Then et = "FILE"
            isFolder = (et = "FOLDER")
            ' Folders: own name must match. SQL InStr on FilePath also hits descendants
            ' of a matching ancestor (Pics, IM-000020, ...); drop those here.
            If isFolder Then
                If Not modPathUtil.FolderLeafMatchesCriteria(path, term1, op, term2, flexible, extraAndTerm) Then GoTo NextRec
            ElseIf flexible Then
                If Not FlexiblePathMatches(path, t1, op, t2) Then GoTo NextRec
                If Len(extraN) > 0 Then
                    If InStr(1, modPathUtil.NormalizeForMatch(path), extraN, vbBinaryCompare) = 0 Then GoTo NextRec
                End If
            End If
            sizeVal = modPathUtil.CoerceStoredSizeMb(rs.Fields("SizeMB").Value)
            If hasSizeFilter Then
                If sizeOp = "OVER" Then
                    If Not (sizeVal >= sizeMb) Then GoTo NextRec
                ElseIf sizeOp = "UNDER" Then
                    If Not (sizeVal <= sizeMb) Then GoTo NextRec
                End If
            End If
            dt = rs.Fields("FileDate").Value
            If Not hits.Exists(path) Then
                hits.Add path, Array(dt, sizeVal, isFolder, 0&, 0&)
            End If
        End If
NextRec:
        rs.MoveNext
    Loop
    rs.Close
End Sub

Private Function FlexiblePathMatches(ByVal path As String, ByVal term1N As String, _
                                     ByVal op As String, ByVal term2N As String) As Boolean
    Dim hay As String
    Dim has1 As Boolean
    Dim has2 As Boolean
    Dim hasTerm2 As Boolean

    hay = modPathUtil.NormalizeForMatch(path)
    has1 = (InStr(1, hay, term1N, vbBinaryCompare) > 0)
    hasTerm2 = (Len(term2N) > 0 And Len(op) > 0)
    If hasTerm2 Then
        has2 = (InStr(1, hay, term2N, vbBinaryCompare) > 0)
        Select Case UCase$(op)
            Case "AND": FlexiblePathMatches = has1 And has2
            Case "OR": FlexiblePathMatches = has1 Or has2
            Case "NOT": FlexiblePathMatches = has1 And (Not has2)
            Case Else: FlexiblePathMatches = has1
        End Select
    Else
        FlexiblePathMatches = has1
    End If
End Function

Private Function BuildAccdbWhere(ByVal term1 As String, ByVal op As String, ByVal term2 As String, _
                                 ByVal includeFiles As Boolean, ByVal includeFolders As Boolean, _
                                 ByVal hasSizeFilter As Boolean, ByVal sizeOp As String, ByVal sizeMb As Double, _
                                 ByVal shareFilter As String, _
                                 Optional ByVal extraAndTerm As String = "") As String
    Dim parts As String
    Dim kw As String
    Dim shareEsc As String
    Dim extra As String

    kw = KeywordClause(term1, op, term2)
    If Len(Trim$(kw)) = 0 Then
        BuildAccdbWhere = vbNullString
        Exit Function
    End If
    parts = "(" & kw & ")"

    extra = Trim$(extraAndTerm)
    If Len(extra) > 0 Then
        parts = parts & " AND (" & PathContainsClause(extra) & ")"
    End If

    If includeFiles And Not includeFolders Then
        parts = parts & " AND (UCase([EntryType])='FILE' OR [EntryType] IS NULL OR Len([EntryType] & '')=0)"
    ElseIf includeFolders And Not includeFiles Then
        parts = parts & " AND UCase([EntryType])='FOLDER'"
    End If

    If hasSizeFilter Then
        ' Loose SQL prefilter (raw column). Authoritative match is CoerceStoredSizeMb post-filter.
        If sizeOp = "OVER" Then
            parts = parts & " AND [SizeMB]>=" & Replace(CStr(sizeMb), ",", ".")
        ElseIf sizeOp = "UNDER" Then
            parts = parts & " AND [SizeMB]<=" & Replace(CStr(sizeMb), ",", ".")
        End If
    End If

    If Len(Trim$(shareFilter)) > 0 Then
        shareEsc = EscapeAccessString(shareFilter)
        parts = parts & " AND InStr(1, [FilePath], '" & shareEsc & "')=1"
    End If

    BuildAccdbWhere = parts
End Function

Private Function KeywordClause(ByVal term1 As String, ByVal op As String, ByVal term2 As String) As String
    Dim c1 As String
    Dim c2 As String
    Dim hasTerm2 As Boolean

    ' InStr is reliable via ACE/ADODB; LIKE %/* wildcard mode varies by ANSI setting.
    c1 = PathContainsClause(term1)
    If Len(c1) = 0 Then
        KeywordClause = vbNullString
        Exit Function
    End If
    hasTerm2 = (Len(Trim$(term2)) > 0 And Len(Trim$(op)) > 0)
    If Not hasTerm2 Then
        KeywordClause = c1
        Exit Function
    End If
    c2 = PathContainsClause(term2)
    If Len(c2) = 0 Then
        KeywordClause = c1
        Exit Function
    End If
    Select Case UCase$(Trim$(op))
        Case "AND"
            KeywordClause = "(" & c1 & " AND " & c2 & ")"
        Case "OR"
            KeywordClause = "(" & c1 & " OR " & c2 & ")"
        Case "NOT"
            KeywordClause = "(" & c1 & " AND NOT (" & c2 & "))"
        Case Else
            KeywordClause = c1
    End Select
End Function

Private Function PathContainsClause(ByVal term As String) As String
    Dim t As String
    t = Trim$(term)
    If Len(t) = 0 Then
        PathContainsClause = vbNullString
        Exit Function
    End If
    ' Case-insensitive path substring match
    PathContainsClause = "InStr(1, LCase([FilePath]), LCase('" & EscapeAccessString(t) & "'))>0"
End Function

Private Function EscapeAccessString(ByVal text As String) As String
    EscapeAccessString = Replace(text, "'", "''")
End Function
Private Function ParseSizeFilter(ByRef sizeOp As String, ByVal sizeValue As Variant, ByRef sizeMb As Double) As Boolean
    Dim opNorm As String
    Dim raw As String

    opNorm = UCase$(Trim$(sizeOp))
    If opNorm = "GREATER THAN" Or opNorm = "GT" Or opNorm = ">" Or opNorm = "OVER" Then
        opNorm = "OVER"
    ElseIf opNorm = "LESS THAN" Or opNorm = "LT" Or opNorm = "<" Or opNorm = "UNDER" Then
        opNorm = "UNDER"
    Else
        ParseSizeFilter = False
        Exit Function
    End If
    sizeOp = opNorm

    raw = Trim$(CStr(sizeValue & ""))
    If Len(raw) = 0 Or Not IsNumeric(raw) Then
        ParseSizeFilter = False
        Exit Function
    End If

    sizeMb = CDbl(raw)
    If sizeMb < 0 Then
        ParseSizeFilter = False
        Exit Function
    End If

    ParseSizeFilter = True
End Function

Private Sub ClearResults()
    Dim ws As Worksheet
    Dim tbl As ListObject

    Set ws = mMap.Dashboard

    mMap.UnprotectDashboard
    mMap.ClearResultsAutoFilter

    On Error Resume Next
    Set tbl = ws.ListObjects(mMap.TableResultsName)
    If Not tbl Is Nothing Then tbl.Delete
    On Error GoTo 0

    mMap.ClearResultsValues

    mMap.EnsureDashboardChromeMerges
End Sub

Private Sub WriteResults(ByVal hits As Object)
    Dim ws As Worksheet
    Dim firstDataRow As Long
    Dim fileCol As Long
    Dim typeCol As Long
    Dim sizeCol As Long
    Dim dateCol As Long
    Dim driveCol As Long
    Dim folderCol As Long
    Dim keys As Variant
    Dim i As Long
    Dim n As Long
    Dim uncPath As String
    Dim fileName As String
    Dim displayName As String
    Dim fileType As String
    Dim sorted() As String
    Dim sortKey() As String
    Dim meta As Variant
    Dim outFile() As Variant
    Dim outType() As Variant
    Dim outSize() As Variant
    Dim outDate() As Variant
    Dim outDrive() As Variant
    Dim outFolder() As Variant
    Dim prevFolderKey As String
    Dim groupFolderUnc As String
    Dim groupFolderDisplay As String
    Dim groupDrive As String
    Dim isFolderHit As Boolean
    Dim fileCount As Long
    Dim folderCount As Long
    Dim totalSizeMb As Double
    Dim tPhase As Double
    Dim outLink() As Variant
    Dim linkCol As Long
    Dim sizeRng As Range
    Dim fileRng As Range
    Dim writeN As Long
    Dim maxRows As Long

    maxRows = mMap.ResultsMaxRows

    Set ws = mMap.Dashboard
    firstDataRow = mMap.ResultsFirstDataRow
    fileCol = mMap.ResultsFileCol
    typeCol = mMap.ResultsTypeCol
    sizeCol = mMap.ResultsSizeCol
    dateCol = mMap.ResultsDateCol
    driveCol = mMap.ResultsDriveCol
    folderCol = mMap.ResultsFolderCol
    linkCol = mMap.ResultsLinkCol

    tPhase = Timer
    mMap.EnsureResultsDisplayGrid
    Debug.Print "SEARCH: EnsureResultsDisplayGrid=" & FormatSeconds(ElapsedSeconds(tPhase))

    tPhase = Timer
    ClearResults
    Debug.Print "SEARCH: ClearResults=" & FormatSeconds(ElapsedSeconds(tPhase))

    n = hits.Count
    If n = 0 Then
        mLastHitCount = 0
        mLastDisplayedCount = 0
        WriteSearchStats 0, 0, 0#
        mMap.ProtectDashboardResults
        Debug.Print "SEARCH: WriteResults no hits after clear"
        Exit Sub
    End If

    writeN = n
    If writeN > maxRows Then
        writeN = maxRows
        Debug.Print "SEARCH: truncating display to " & CStr(maxRows) & " of " & CStr(n) & " hits"
    End If
    Debug.Print "SEARCH: writeN=" & CStr(writeN) & " (prebuilt merges/stripes; values+links only)"

    tPhase = Timer
    keys = hits.Keys
    ReDim sorted(1 To n)
    ReDim sortKey(1 To n)
    If n = 1 Then
        ' Scripting.Dictionary.Keys is a scalar for 1 item on some hosts, an array on others
        If IsArray(keys) Then
            sorted(1) = CStr(keys(LBound(keys)))
        Else
            sorted(1) = CStr(keys)
        End If
        sortKey(1) = ResultSortKey(sorted(1), hits(sorted(1)))
    Else
        For i = 0 To n - 1
            sorted(i + 1) = CStr(keys(i))
            sortKey(i + 1) = ResultSortKey(sorted(i + 1), hits(sorted(i + 1)))
        Next i
        QuickSort sorted, sortKey, 1, n
    End If
    Debug.Print "SEARCH: sort hits=" & CStr(n) & " " & FormatSeconds(ElapsedSeconds(tPhase))

    tPhase = Timer
    ReDim outFile(1 To writeN, 1 To 1)
    ReDim outType(1 To writeN, 1 To 1)
    ReDim outSize(1 To writeN, 1 To 1)
    ReDim outDate(1 To writeN, 1 To 1)
    ReDim outDrive(1 To writeN, 1 To 1)
    ReDim outFolder(1 To writeN, 1 To 1)
    ReDim outLink(1 To writeN, 1 To 1)

    prevFolderKey = vbNullString
    fileCount = 0
    folderCount = 0
    totalSizeMb = 0#
    For i = 1 To n
        uncPath = sorted(i)
        meta = hits(uncPath)
        isFolderHit = (IsArray(meta) And UBound(meta) >= 2 And CBool(meta(2)))

        If isFolderHit Then
            folderCount = folderCount + 1
            groupFolderUnc = uncPath
            fileName = modPathUtil.FileNameOnly(uncPath)
            fileType = "FOLDER"
            displayName = "\" & fileName & "\"
            If i <= writeN Then
                outFile(i, 1) = displayName
                outLink(i, 1) = uncPath ' double-click opens this folder
                outType(i, 1) = fileType
                outSize(i, 1) = 0.01
                If IsArray(meta) Then
                    ' Recursive disk total of all files under this folder (indexed SizeMB).
                    outSize(i, 1) = modPathUtil.CoerceStoredSizeMb(meta(1))
                End If
                outDate(i, 1) = Empty
                If IsArray(meta) Then
                    If Not IsEmpty(meta(0)) Then
                        On Error Resume Next
                        outDate(i, 1) = DateValue(CDate(meta(0)))
                        On Error GoTo 0
                    End If
                End If
            End If
            If IsArray(meta) Then
                totalSizeMb = totalSizeMb + CDbl(modPathUtil.CoerceStoredSizeMb(meta(1)))
            Else
                totalSizeMb = totalSizeMb + 0.01
            End If
        Else
            fileCount = fileCount + 1
            fileName = modPathUtil.FileNameOnly(uncPath)
            fileType = modPathUtil.FileExtension(fileName)
            displayName = AppendTypeToFileName(fileName, fileType)
            groupFolderUnc = modPathUtil.ParentFolderPath(uncPath)
            If Right$(groupFolderUnc, 1) = "\" Then groupFolderUnc = Left$(groupFolderUnc, Len(groupFolderUnc) - 1)

            If i <= writeN Then
                outFile(i, 1) = displayName
                outLink(i, 1) = groupFolderUnc ' double-click opens containing folder
                outType(i, 1) = fileType
                outSize(i, 1) = 0.01
                outDate(i, 1) = Empty
                If IsArray(meta) Then
                    outSize(i, 1) = modPathUtil.CoerceStoredSizeMb(meta(1))
                    If Not IsEmpty(meta(0)) Then
                        On Error Resume Next
                        outDate(i, 1) = DateValue(CDate(meta(0)))
                        On Error GoTo 0
                    End If
                End If
            End If
            If IsArray(meta) Then
                totalSizeMb = totalSizeMb + CDbl(modPathUtil.CoerceStoredSizeMb(meta(1)))
            Else
                totalSizeMb = totalSizeMb + 0.01
            End If
        End If

        If i <= writeN Then
            If StrComp(groupFolderUnc, prevFolderKey, vbTextCompare) = 0 Then
                outDrive(i, 1) = Empty
                outFolder(i, 1) = Empty
            Else
                ' Resolve drive label only when the folder group changes (was per-row Shell/WNet)
                groupDrive = modIndexCache.FriendlyLabelForShare(groupFolderUnc)
                groupFolderDisplay = modPathUtil.ToDisplayPath(groupFolderUnc)
                If isFolderHit Then
                    If Right$(groupFolderDisplay, 1) <> "\" Then groupFolderDisplay = groupFolderDisplay & "\"
                End If
                outDrive(i, 1) = groupDrive
                outFolder(i, 1) = groupFolderDisplay
                prevFolderKey = groupFolderUnc
            End If
        End If
    Next i
    Debug.Print "SEARCH: build output arrays=" & FormatSeconds(ElapsedSeconds(tPhase))

    tPhase = Timer
    ' Batch dump only — no Hyperlinks.Add (double-click opens folder via hidden UNC col)
    Set sizeRng = ws.Range(ws.Cells(firstDataRow, sizeCol), ws.Cells(firstDataRow + writeN - 1, sizeCol))
    Set fileRng = ws.Range(ws.Cells(firstDataRow, fileCol), ws.Cells(firstDataRow + writeN - 1, fileCol))
    sizeRng.NumberFormat = "0.00"
    ' Do not format A8:An as a block (collapses A:K row merges). File cells are "@" from the grid.
    fileRng.Value = outFile
    ws.Range(ws.Cells(firstDataRow, typeCol), ws.Cells(firstDataRow + writeN - 1, typeCol)).Value2 = outType
    sizeRng.Value2 = outSize
    ws.Range(ws.Cells(firstDataRow, dateCol), ws.Cells(firstDataRow + writeN - 1, dateCol)).NumberFormat = "yyyy-mmm-dd"
    ws.Range(ws.Cells(firstDataRow, dateCol), ws.Cells(firstDataRow + writeN - 1, dateCol)).Value2 = outDate
    ws.Range(ws.Cells(firstDataRow, driveCol), ws.Cells(firstDataRow + writeN - 1, driveCol)).Value2 = outDrive
    ws.Range(ws.Cells(firstDataRow, folderCol), ws.Cells(firstDataRow + writeN - 1, folderCol)).Value2 = outFolder
    ws.Range(ws.Cells(firstDataRow, linkCol), ws.Cells(firstDataRow + writeN - 1, linkCol)).Value2 = outLink

    sizeRng.NumberFormat = "0.00"
    sizeRng.HorizontalAlignment = xlRight
    ws.Range(ws.Cells(firstDataRow, typeCol), ws.Cells(firstDataRow + writeN - 1, typeCol)).HorizontalAlignment = xlCenter
    ' Do not set File Font/Alignment on A8:An as a block — collapses per-row A:K merges into A8:An
    ' Folder Path / Drive Address: plain text only (blank repeats are intentional grouping)
    With ws.Range(ws.Cells(firstDataRow, driveCol), ws.Cells(firstDataRow + writeN - 1, folderCol))
        .Font.ColorIndex = xlColorIndexAutomatic
        .Font.Underline = xlUnderlineStyleNone
    End With
    Debug.Print "SEARCH: sheet dump+formats=" & FormatSeconds(ElapsedSeconds(tPhase)) & _
                " sizeSample=" & CStr(outSize(1, 1)) & " (no hyperlink COM)"

    WriteSearchStats fileCount, folderCount, Round(totalSizeMb, 2)

    mLastHitCount = n
    mLastDisplayedCount = writeN

    mMap.EnsureDashboardChromeMerges
    mMap.ProtectDashboardResults
End Sub

Private Sub WriteSearchStats(ByVal fileCount As Long, ByVal folderCount As Long, ByVal totalSizeMb As Double)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = mMap.Dashboard
    ' Counts/sizes — never leave these cells (column M) as Percentage
    ws.Range(mMap.StatFilesValueAddress).NumberFormat = "#,##0"
    ws.Range(mMap.StatFoldersValueAddress).NumberFormat = "#,##0"
    ws.Range(mMap.StatSizeValueAddress).NumberFormat = "#,##0.00"
    ws.Range(mMap.StatFilesValueAddress).Value = fileCount
    ws.Range(mMap.StatFoldersValueAddress).Value = folderCount
    ws.Range(mMap.StatSizeValueAddress).Value = totalSizeMb
    On Error GoTo 0
End Sub

Private Function AppendTypeToFileName(ByVal fileName As String, ByVal fileType As String) As String
    Dim n As String
    Dim t As String

    n = Trim$(fileName)
    t = UCase$(Trim$(fileType))
    If Len(n) = 0 Then
        AppendTypeToFileName = n
        Exit Function
    End If
    If Len(t) = 0 Then
        AppendTypeToFileName = n
        Exit Function
    End If
    If Right$(UCase$(n), Len(t) + 1) = "." & t Then
        AppendTypeToFileName = n
    Else
        AppendTypeToFileName = n & "." & t
    End If
End Function

Private Function ResultSortKey(ByVal uncPath As String, ByVal meta As Variant) As String
    Dim folderUnc As String
    Dim namePart As String
    Dim isFolderHit As Boolean

    isFolderHit = (IsArray(meta) And UBound(meta) >= 2 And CBool(meta(2)))
    If isFolderHit Then
        folderUnc = uncPath
        namePart = " " ' folder summary rows sort before files in same path
    Else
        folderUnc = modPathUtil.ParentFolderPath(uncPath)
        If Right$(folderUnc, 1) = "\" Then folderUnc = Left$(folderUnc, Len(folderUnc) - 1)
        namePart = LCase$(modPathUtil.FileNameOnly(uncPath))
    End If
    ResultSortKey = LCase$(modPathUtil.ToDisplayPath(folderUnc)) & vbNullChar & namePart
End Function

Private Sub QuickSort(ByRef values() As String, ByRef keys() As String, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long
    Dim j As Long
    Dim pivot As String
    Dim tmpV As String
    Dim tmpK As String

    i = lo
    j = hi
    pivot = keys((lo + hi) \ 2)
    Do While i <= j
        Do While StrComp(keys(i), pivot, vbBinaryCompare) < 0
            i = i + 1
        Loop
        Do While StrComp(keys(j), pivot, vbBinaryCompare) > 0
            j = j - 1
        Loop
        If i <= j Then
            tmpV = values(i)
            values(i) = values(j)
            values(j) = tmpV
            tmpK = keys(i)
            keys(i) = keys(j)
            keys(j) = tmpK
            i = i + 1
            j = j - 1
        End If
    Loop
    If lo < j Then QuickSort values, keys, lo, j
    If i < hi Then QuickSort values, keys, i, hi
End Sub

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
