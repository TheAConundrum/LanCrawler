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

' Searches Database!tblFiles in memory:
' - File hits: match against file name only
' - Folder hits (optional, default on): match folder names, one collapsed row
'   with recursive file/folder counts (not every file under that folder)

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

    tAll = Timer
    calcPrev = Application.Calculation
    On Error GoTo Fail
    mApp.BeginFast
    Application.Calculation = xlCalculationManual

    Set hits = CreateObject("Scripting.Dictionary")
    hits.CompareMode = 1 ' TextCompare

    Debug.Print "========== SEARCH TIMING START =========="
    Debug.Print "SEARCH: cache already loaded before get? " & CStr(modIndexCache.IsIndexLoaded) & _
                " rows=" & CStr(modIndexCache.IndexRowCount)

    tPhase = Timer
    modIndexCache.EnsureLoaded ActiveWorkbook
    secGet = ElapsedSeconds(tPhase)
    cacheHit = modIndexCache.LastGetWasCacheHit
    rowCount = modIndexCache.IndexRowCount

    If cacheHit Then
        Debug.Print "SEARCH: EnsureLoaded HIT rows=" & CStr(rowCount) & _
                    " elapsed=" & FormatSeconds(secGet)
    Else
        Debug.Print "SEARCH: EnsureLoaded MISS rows=" & CStr(rowCount) & _
                    " elapsed=" & FormatSeconds(secGet) & _
                    " load=" & FormatSeconds(modIndexCache.LastLoadSeconds)
    End If

    If rowCount = 0 Then
        ClearResults
        WriteSearchStats 0, 0, 0#
        mMap.ProtectDashboardResults
        Application.Calculation = calcPrev
        mApp.Restore
        Debug.Print "SEARCH: aborted empty database total=" & FormatSeconds(ElapsedSeconds(tAll))
        MsgBox "Database is empty. Ingest a folder first.", vbExclamation, "LAN Search - Dashboard"
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

    Debug.Print "SEARCH SUMMARY: get=" & FormatSeconds(secGet) & _
                " scan=" & FormatSeconds(secScan) & _
                " write=" & FormatSeconds(secWrite) & _
                " TOTAL=" & FormatSeconds(ElapsedSeconds(tAll)) & _
                " cacheHit=" & CStr(cacheHit) & " rows=" & CStr(rowCount) & " hits=" & CStr(hits.Count)
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
    Exit Sub

Fail:
    Dim errDesc As String
    errDesc = Err.Description
    On Error Resume Next
    Application.Calculation = calcPrev
    mApp.Restore
    Debug.Print "SEARCH FAILED: " & errDesc & " after " & FormatSeconds(ElapsedSeconds(tAll))
    On Error GoTo 0
    MsgBox "Search failed: " & errDesc, vbCritical, "LAN Search - Dashboard"
End Sub

' Uses modIndexCache parallel arrays — no per-search Variant copy / re-normalize.
Private Sub SearchIndex(ByVal hits As Object, ByRef anyCriteria As Boolean)
    Dim row As Long
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
    Dim folderHits As Object
    Dim tScan As Double
    Dim tCommit As Double
    Dim fileHitsBefore As Long
    Dim shareFilter As String

    flexible = mMap.FlexibleMatch
    foldersMode = mMap.SearchFoldersMode
    includeFiles = (foldersMode <> "Only")
    includeFolders = (foldersMode = "Yes" Or foldersMode = "Only")
    shareFilter = modIndexCache.ResolveDriveShareFilter(mMap.DriveFilterLabel)

    Debug.Print "SEARCH: foldersMode=" & foldersMode & _
                " includeFiles=" & CStr(includeFiles) & _
                " includeFolders=" & CStr(includeFolders) & _
                " flexible=" & CStr(flexible) & _
                " drive=" & mMap.DriveFilterLabel & _
                " shareFilter=" & shareFilter & _
                " n=" & CStr(modIndexCache.IndexRowCount)

    For row = mMap.CriteriaFirstRow To mMap.CriteriaLastRow
        term1 = Trim$(CStr(mMap.Dashboard.Cells(row, mMap.CriteriaColTerm1).Value & ""))
        If Len(term1) = 0 Then Exit For
        anyCriteria = True

        op = UCase$(Trim$(CStr(mMap.Dashboard.Cells(row, mMap.CriteriaColOp).Value & "")))
        term2 = Trim$(CStr(mMap.Dashboard.Cells(row, mMap.CriteriaColTerm2).Value & ""))
        sizeOp = Trim$(CStr(mMap.Dashboard.Cells(row, mMap.CriteriaColSizeOp).Value & ""))
        hasSizeFilter = ParseSizeFilter( _
            sizeOp, _
            mMap.Dashboard.Cells(row, mMap.CriteriaColSizeMb).Value, _
            sizeMb)

        If op = "NOT" And Len(term2) = 0 Then
            Err.Raise vbObjectError + 1, , "NOT requires a second Key Word (row " & CStr(row) & ")."
        End If
        If Len(sizeOp) > 0 And Not hasSizeFilter Then
            Err.Raise vbObjectError + 2, , "Size filter incomplete (row " & CStr(row) & ")."
        End If

        Set folderHits = CreateObject("Scripting.Dictionary")
        folderHits.CompareMode = 1
        fileHitsBefore = hits.Count

        tScan = Timer
        modIndexCache.CollectNameHits term1, op, term2, flexible, _
            includeFiles, includeFolders, hasSizeFilter, sizeOp, sizeMb, _
            hits, folderHits, shareFilter
        Debug.Print "SEARCH: criteria row " & CStr(row) & " scan=" & FormatSeconds(ElapsedSeconds(tScan)) & _
                    " term1=" & term1 & " fileHits+" & CStr(hits.Count - fileHitsBefore) & _
                    " folderNameHits=" & CStr(folderHits.Count)

        If includeFolders Then
            tCommit = Timer
            modIndexCache.CommitFolderHitsFast hits, folderHits, hasSizeFilter, sizeOp, sizeMb
            Debug.Print "SEARCH: criteria row " & CStr(row) & " CommitFolderHits=" & _
                        FormatSeconds(ElapsedSeconds(tCommit)) & _
                        " folderHits=" & CStr(folderHits.Count) & " (precomputed counts)"
        End If
    Next row
End Sub

Private Function ParseSizeFilter(ByRef sizeOp As String, ByVal sizeValue As Variant, ByRef sizeMb As Double) As Boolean
    Dim opNorm As String
    Dim raw As String

    opNorm = Trim$(sizeOp)
    If opNorm = "GREATER THAN" Or opNorm = "GT" Then opNorm = ">"
    If opNorm = "LESS THAN" Or opNorm = "LT" Then opNorm = "<"

    If opNorm <> "<" And opNorm <> ">" Then
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
    Dim firstDataRow As Long
    Dim lastCol As Long
    Dim linkCol As Long
    Dim clearTo As Long
    Dim maxRows As Long

    Set ws = mMap.Dashboard
    firstDataRow = mMap.ResultsFirstDataRow
    lastCol = mMap.ResultsLastCol
    linkCol = mMap.ResultsLinkCol
    maxRows = mMap.ResultsMaxRows
    If firstDataRow <= mMap.ResultsHeaderRow Then firstDataRow = mMap.ResultsHeaderRow + 1

    mMap.UnprotectDashboard

    On Error Resume Next
    Set tbl = ws.ListObjects(mMap.TableResultsName)
    If Not tbl Is Nothing Then tbl.Delete
    On Error GoTo 0

    clearTo = firstDataRow + maxRows - 1

    With ws.Range(ws.Cells(firstDataRow, 1), ws.Cells(clearTo, lastCol))
        .ClearContents
        On Error Resume Next
        .Hyperlinks.Delete
        .ClearHyperlinks
        On Error GoTo 0
    End With
    ws.Range(ws.Cells(firstDataRow, linkCol), ws.Cells(clearTo, linkCol)).ClearContents

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
        sorted(1) = CStr(keys)
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
            displayName = modPathUtil.FileNameOnly(uncPath) & " (" & _
                          Format$(CLng(meta(3)), "#,##0") & " files, " & _
                          Format$(CLng(meta(4)), "#,##0") & " folders)"
            If i <= writeN Then
                outFile(i, 1) = displayName
                outLink(i, 1) = uncPath ' double-click opens this folder
                outType(i, 1) = "FOLDER"
                outSize(i, 1) = Round(CDbl(meta(1) & 0), 2)
                outDate(i, 1) = Empty
                If Not IsEmpty(meta(0)) Then
                    On Error Resume Next
                    outDate(i, 1) = DateValue(CDate(meta(0)))
                    On Error GoTo 0
                End If
            End If
            totalSizeMb = totalSizeMb + Round(CDbl(meta(1) & 0), 2)
        Else
            fileCount = fileCount + 1
            fileName = modPathUtil.FileNameOnly(uncPath)
            displayName = fileName
            fileType = modPathUtil.FileExtension(fileName)
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
                totalSizeMb = totalSizeMb + CDbl(modPathUtil.CoerceStoredSizeMb(meta(1)) & 0)
            Else
                totalSizeMb = totalSizeMb + 0.01
            End If
        End If

        If i <= writeN Then
            groupDrive = modIndexCache.FriendlyLabelForShare(groupFolderUnc)
            groupFolderDisplay = modPathUtil.ToDisplayPath(groupFolderUnc)
            If StrComp(groupFolderUnc, prevFolderKey, vbTextCompare) = 0 Then
                outDrive(i, 1) = Empty
                outFolder(i, 1) = Empty
            Else
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

    fileRng.Value2 = outFile
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
    With fileRng.Font
        .Color = RGB(5, 99, 193)
        .Underline = xlUnderlineStyleSingle
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
