VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFileIndexer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Indexes LAN folders into Database!tblFiles (UNC, FileDate, SizeMB, EntryType).
' EntryType = FILE | FOLDER. Default mode: ReplaceRoot.
' Folder SizeMB = recursive disk total of ALL files under that folder (incl. non-indexed
' extensions), accumulated during crawl — no extra disk walk.
' Updates Ingestion!tblIngested with counts and TotalSizeMB.

Private mMap As clsSheetMap
Private mApp As clsExcelAppState
Private mBuffer() As Variant ' col1 path, col2 FileDate, col3 SizeMB, col4 EntryType
Private mCount As Long       ' buffer rows (files + folders)
Private mFileCount As Long   ' FILE rows only (tblIngested FileCount)
Private mCapacity As Long
Private mFolderCount As Long
Private mTotalBytes As Double
Private mSkipLog As Object
Private mStage As String
Private mFolderBytes As Object ' Scripting.Dictionary: folder UNC -> Double bytes (disk total)
Private mCrawlRootUnc As String

Private Const INITIAL_CAPACITY As Long = 5000
Private Const BYTES_PER_MB As Double = 1048576#
Private Const COLS_FILES As Long = 4
Private Const ENTRY_FILE As String = "FILE"
Private Const ENTRY_FOLDER As String = "FOLDER"

Public Sub Init(ByVal map As clsSheetMap, ByVal appState As clsExcelAppState)
    Set mMap = map
    Set mApp = appState
End Sub

Public Function RunFromIngestion() As Long
    Dim mode As String
    Dim chosen As String

    mode = mMap.IngestMode
    Select Case mode
        Case "RebuildAll"
            ' Rebuild uses Links already Eaten — no folder picker
            RunFromIngestion = RebuildAllFromIngestedList()
        Case "Append"
            chosen = modPathUtil.BrowseForFolder("Select a folder to append to the index")
            If Len(chosen) = 0 Then
                RunFromIngestion = 0
                Exit Function
            End If
            RunFromIngestion = IndexRoot(chosen, "Append")
        Case Else
            ' ReplaceRoot (default): pick folder in Explorer
            chosen = modPathUtil.BrowseForFolder("Select a folder to ingest")
            If Len(chosen) = 0 Then
                RunFromIngestion = 0
                Exit Function
            End If
            RunFromIngestion = IndexRoot(chosen, "ReplaceRoot")
    End Select
End Function

Public Function IndexRoot(ByVal rootPath As String, ByVal mode As String) As Long
    Dim uncRoot As String
    Dim fso As Object
    Dim skipCount As Long
    Dim crawlStart As String

    On Error GoTo Fail

    SetStage "BeginFast"
    mApp.BeginFast
    mMap.UnprotectWorkbookForWrite
    Set mSkipLog = CreateObject("Scripting.Dictionary")

    rootPath = Trim$(rootPath)
    If Len(rootPath) = 0 Then
        Err.Raise vbObjectError + 1, "clsFileIndexer", "No folder was selected."
    End If

    SetStage "Resolve UNC from " & rootPath
    uncRoot = modPathUtil.ToUncPath(rootPath, mMap.UncRootOverride)
    If Right$(uncRoot, 1) = "\" Then uncRoot = Left$(uncRoot, Len(uncRoot) - 1)
    Debug.Print "INGEST: selected=[" & rootPath & "] unc=[" & uncRoot & "] mode=[" & mode & "]"

    SetStage "Validate folder exists"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(uncRoot) Then
        crawlStart = uncRoot
    ElseIf fso.FolderExists(rootPath) Then
        crawlStart = rootPath
    Else
        Err.Raise vbObjectError + 2, "clsFileIndexer", "Folder not found: " & rootPath
    End If

    SetStage "Crawl " & crawlStart
    ResetBuffer
    CrawlIterative crawlStart
    FinalizeFolderSizes
    Debug.Print "INGEST: crawl done files=" & mFileCount & " folders=" & mFolderCount & _
                " entries=" & mCount & " indexedBytes=" & mTotalBytes

    SetStage "Write Database (tblFiles) mode=" & mode
    ApplyIndexWrite uncRoot, mode

    SetStage "Update Links already Eaten (tblIngested)"
    UpsertIngestedList uncRoot, Date

    skipCount = mSkipLog.Count
    IndexRoot = mFileCount
    mMap.ProtectWorkbookUi
    mApp.Restore
    SetStage "Complete"

    If skipCount > 0 Then
        DumpSkipLog
        MsgBox "Indexed " & Format$(mFileCount, "#,##0") & " file(s), " & Format$(mFolderCount, "#,##0") & " folder(s), " & _
               Format$(mTotalBytes / BYTES_PER_MB, "#,##0.00") & " MB under:" & vbCrLf & uncRoot & vbCrLf & vbCrLf & _
               "Skipped " & CStr(skipCount) & " inaccessible folder(s). See Immediate Window (Ctrl+G).", _
               vbInformation, "LAN Search - Ingestion"
    Else
        MsgBox "Indexed " & Format$(mFileCount, "#,##0") & " file(s), " & Format$(mFolderCount, "#,##0") & " folder(s), " & _
               Format$(mTotalBytes / BYTES_PER_MB, "#,##0.00") & " MB under:" & vbCrLf & uncRoot, _
               vbInformation, "LAN Search - Ingestion"
    End If
    Exit Function

Fail:
    ReportFailure "IndexRoot"
    IndexRoot = -1
End Function

Public Function RebuildAllFromIngestedList() As Long
    Dim tbl As ListObject
    Dim r As Long
    Dim root As String
    Dim total As Long
    Dim n As Long
    Dim roots() As String
    Dim rootCount As Long

    On Error GoTo Fail
    mApp.BeginFast
    mMap.UnprotectWorkbookForWrite

    EnsureFilesTableEmpty
    total = 0

    Set tbl = GetOrCreateIngestedTable()
    If tbl.DataBodyRange Is Nothing Then
        mMap.ProtectWorkbookUi
        mApp.Restore
        MsgBox "Links already Eaten is empty. Add roots via ReplaceRoot first.", vbExclamation, "LAN Search - RebuildAll"
        RebuildAllFromIngestedList = 0
        Exit Function
    End If

    rootCount = tbl.DataBodyRange.Rows.Count
    ReDim roots(1 To rootCount)
    For r = 1 To rootCount
        roots(r) = CStr(tbl.DataBodyRange.Cells(r, 1).Value & "")
    Next r

    mApp.Restore

    For r = 1 To rootCount
        root = Trim$(roots(r))
        If Len(root) > 0 Then
            n = IndexRootSilent(root, "ReplaceRoot")
            If n > 0 Then total = total + n
        End If
    Next r

    mMap.ProtectWorkbookUi
    RebuildAllFromIngestedList = total
    MsgBox "RebuildAll complete. Files indexed this run (sum of roots): " & CStr(total), vbInformation, "LAN Search - RebuildAll"
    Exit Function

Fail:
    ReportFailure "RebuildAll"
    RebuildAllFromIngestedList = -1
End Function

Private Function IndexRootSilent(ByVal rootPath As String, ByVal mode As String) As Long
    Dim uncRoot As String
    Dim fso As Object
    Dim crawlStart As String

    On Error GoTo Fail
    mApp.BeginFast
    mMap.UnprotectWorkbookForWrite
    Set mSkipLog = CreateObject("Scripting.Dictionary")

    uncRoot = modPathUtil.ToUncPath(rootPath, mMap.UncRootOverride)
    If Right$(uncRoot, 1) = "\" Then uncRoot = Left$(uncRoot, Len(uncRoot) - 1)

    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(uncRoot) Then
        crawlStart = uncRoot
    ElseIf fso.FolderExists(rootPath) Then
        crawlStart = rootPath
    Else
        IndexRootSilent = 0
        mApp.Restore
        Exit Function
    End If

    ResetBuffer
    CrawlIterative crawlStart
    FinalizeFolderSizes
    ApplyIndexWrite uncRoot, mode
    UpsertIngestedList uncRoot, Date
    IndexRootSilent = mFileCount
    mApp.Restore
    Exit Function

Fail:
    ReportFailure "IndexRootSilent"
    IndexRootSilent = -1
End Function

Public Sub RefreshIngestedListFromDatabase()
    Dim tbl As ListObject
    Dim files As ListObject
    Dim r As Long
    Dim root As String
    Dim found As Boolean
    Dim cell As Range

    On Error GoTo Fail
    mApp.BeginFast
    mMap.UnprotectWorkbookForWrite

    Set tbl = GetOrCreateIngestedTable()
    Set files = GetFilesTable()

    If tbl.DataBodyRange Is Nothing Then
        mMap.ProtectWorkbookUi
        mApp.Restore
        MsgBox "Links already Eaten is empty.", vbInformation, "LAN Search"
        Exit Sub
    End If

    For r = tbl.DataBodyRange.Rows.Count To 1 Step -1
        root = CStr(tbl.DataBodyRange.Cells(r, 1).Value & "")
        found = False
        If Not files Is Nothing Then
            If Not files.DataBodyRange Is Nothing Then
                For Each cell In files.ListColumns(1).DataBodyRange.Cells
                    If modPathUtil.PathStartsWithRoot(CStr(cell.Value & ""), root) Then
                        found = True
                        Exit For
                    End If
                Next cell
            End If
        End If
        If Not found Then tbl.ListRows(r).Delete
    Next r

    SortIngestedTable
    RefreshIngestedSummary
    mMap.ProtectWorkbookUi
    mApp.Restore
    MsgBox "Ingested list refreshed.", vbInformation, "LAN Search"
    Exit Sub

Fail:
    ReportFailure "RefreshIngestedList"
End Sub

Private Sub CrawlIterative(ByVal startFolder As String)
    Dim fso As Object
    Dim stack As Collection
    Dim folderPath As String
    Dim folder As Object
    Dim subFolder As Object
    Dim file As Object
    Dim uncFile As String
    Dim uncFolder As String
    Dim foldersDone As Long
    Dim sizeBytes As Double

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set stack = New Collection
    stack.Add startFolder

    ' Crawl root UNC (no trailing slash) — used when bubbling folder disk totals.
    mCrawlRootUnc = modPathUtil.ToUncPath(startFolder, mMap.UncRootOverride)
    If Right$(mCrawlRootUnc, 1) = "\" Then mCrawlRootUnc = Left$(mCrawlRootUnc, Len(mCrawlRootUnc) - 1)

    Do While stack.Count > 0
        folderPath = CStr(stack(stack.Count))
        stack.Remove stack.Count

        On Error Resume Next
        Set folder = fso.GetFolder(folderPath)
        If Err.Number <> 0 Then
            LogSkip folderPath, Err.Description
            Err.Clear
            On Error GoTo 0
            GoTo NextFolder
        End If
        On Error GoTo 0

        If modPathUtil.IsJunkFolderName(folder.Name) Then GoTo NextFolder

        uncFolder = modPathUtil.ToUncPath(folder.Path, mMap.UncRootOverride)
        If Right$(uncFolder, 1) = "\" Then uncFolder = Left$(uncFolder, Len(uncFolder) - 1)

        mFolderCount = mFolderCount + 1
        AddFolderEntry uncFolder, folder

        On Error Resume Next
        For Each file In folder.Files
            If Err.Number <> 0 Then
                LogSkip folderPath & "\*", Err.Description
                Err.Clear
                Exit For
            End If

            ' Disk total for folder rows: every file we can size (indexed or not).
            sizeBytes = 0
            Err.Clear
            sizeBytes = CDbl(file.Size)
            If Err.Number <> 0 Then
                sizeBytes = 0
                Err.Clear
            End If
            If sizeBytes > 0 Then AddFolderTreeBytes uncFolder, sizeBytes

            If modPathUtil.IsJunkFileName(file.Name) Then GoTo ContinueFile
            If Not modPathUtil.IsAllowedIndexExtension(file.Name) Then GoTo ContinueFile

            uncFile = modPathUtil.ToUncPath(file.Path, mMap.UncRootOverride)
            AddFile uncFile, file
ContinueFile:
        Next file
        Err.Clear
        On Error GoTo 0

        On Error Resume Next
        For Each subFolder In folder.SubFolders
            If Err.Number <> 0 Then
                LogSkip folderPath, Err.Description
                Err.Clear
                Exit For
            End If
            If Not modPathUtil.IsJunkFolderName(subFolder.Name) Then
                stack.Add subFolder.Path
            End If
        Next subFolder
        Err.Clear
        On Error GoTo 0

        foldersDone = foldersDone + 1
        If foldersDone Mod 25 = 0 Then
            mApp.SetStatus "Crawling... folders=" & CStr(mFolderCount) & " files=" & CStr(mFileCount)
            DoEvents
        End If
NextFolder:
    Loop
End Sub

' Add file bytes to this folder and every ancestor up to the crawl root (recursive total).
Private Sub AddFolderTreeBytes(ByVal folderUnc As String, ByVal sizeBytes As Double)
    Dim p As String
    Dim parent As String

    If sizeBytes <= 0 Then Exit Sub
    If mFolderBytes Is Nothing Then Set mFolderBytes = CreateObject("Scripting.Dictionary")

    p = folderUnc
    Do While Len(p) > 0
        If mFolderBytes.Exists(p) Then
            mFolderBytes(p) = CDbl(mFolderBytes(p)) + sizeBytes
        Else
            mFolderBytes.Add p, sizeBytes
        End If

        If Len(mCrawlRootUnc) > 0 Then
            If StrComp(p, mCrawlRootUnc, vbTextCompare) = 0 Then Exit Do
        End If

        parent = modPathUtil.ParentFolderPath(p)
        If Right$(parent, 1) = "\" Then parent = Left$(parent, Len(parent) - 1)
        If Len(parent) = 0 Then Exit Do
        If StrComp(parent, p, vbTextCompare) = 0 Then Exit Do

        ' Do not bubble above the crawl root.
        If Len(mCrawlRootUnc) > 0 Then
            If StrComp(parent, mCrawlRootUnc, vbTextCompare) <> 0 Then
                If Not modPathUtil.PathStartsWithRoot(parent, mCrawlRootUnc) Then Exit Do
            End If
        End If

        p = parent
    Loop
End Sub

' Write recursive disk SizeMB onto each FOLDER buffer row after crawl.
Private Sub FinalizeFolderSizes()
    Dim i As Long
    Dim p As String
    Dim bytes As Double

    If mCount = 0 Then Exit Sub
    If mFolderBytes Is Nothing Then Exit Sub

    For i = 1 To mCount
        If UCase$(Trim$(CStr(mBuffer(i, 4) & ""))) = ENTRY_FOLDER Then
            p = CStr(mBuffer(i, 1) & "")
            If mFolderBytes.Exists(p) Then
                bytes = CDbl(mFolderBytes(p))
                mBuffer(i, 3) = modPathUtil.BytesToSizeMb(bytes)
            Else
                mBuffer(i, 3) = 0.01
            End If
        End If
    Next i
End Sub

Private Sub AddFile(ByVal uncPath As String, ByVal file As Object)
    Dim created As Variant
    Dim sizeBytes As Double

    EnsureCapacity mCount + 1
    mCount = mCount + 1
    mBuffer(mCount, 1) = uncPath

    On Error Resume Next
    created = file.DateCreated
    If Err.Number <> 0 Or IsEmpty(created) Then
        Err.Clear
        created = file.DateLastModified
    End If
    If Err.Number <> 0 Then
        mBuffer(mCount, 2) = Empty
        Err.Clear
    Else
        ' Date only (no time)
        mBuffer(mCount, 2) = DateValue(CDate(created))
    End If

    sizeBytes = CDbl(file.Size)
    If Err.Number <> 0 Then
        sizeBytes = 0
        Err.Clear
    End If
    On Error GoTo 0

    mBuffer(mCount, 3) = modPathUtil.BytesToSizeMb(sizeBytes)
    mBuffer(mCount, 4) = ENTRY_FILE
    mTotalBytes = mTotalBytes + sizeBytes
    mFileCount = mFileCount + 1
End Sub

Private Sub AddFolderEntry(ByVal uncPath As String, ByVal folder As Object)
    Dim created As Variant
    Dim p As String

    p = Trim$(uncPath)
    If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    If Len(p) = 0 Then Exit Sub

    EnsureCapacity mCount + 1
    mCount = mCount + 1
    mBuffer(mCount, 1) = p
    mBuffer(mCount, 3) = 0.01
    mBuffer(mCount, 4) = ENTRY_FOLDER

    On Error Resume Next
    created = folder.DateCreated
    If Err.Number <> 0 Or IsEmpty(created) Then
        Err.Clear
        created = folder.DateLastModified
    End If
    If Err.Number <> 0 Then
        mBuffer(mCount, 2) = Empty
        Err.Clear
    Else
        mBuffer(mCount, 2) = DateValue(CDate(created))
    End If
    On Error GoTo 0
End Sub

Private Sub ResetBuffer()
    mCount = 0
    mFileCount = 0
    mFolderCount = 0
    mTotalBytes = 0
    mCrawlRootUnc = vbNullString
    Set mFolderBytes = CreateObject("Scripting.Dictionary")
    mFolderBytes.CompareMode = 1 ' TextCompare
    mCapacity = INITIAL_CAPACITY
    ReDim mBuffer(1 To mCapacity, 1 To COLS_FILES)
End Sub

Private Sub EnsureCapacity(ByVal needed As Long)
    Dim newCap As Long
    Dim newBuf() As Variant
    Dim i As Long
    Dim c As Long

    If needed <= mCapacity Then Exit Sub

    newCap = mCapacity
    Do While newCap < needed
        newCap = newCap * 2
    Loop

    ReDim newBuf(1 To newCap, 1 To COLS_FILES)
    For i = 1 To mCount
        For c = 1 To COLS_FILES
            newBuf(i, c) = mBuffer(i, c)
        Next c
    Next i
    Erase mBuffer
    mBuffer = newBuf
    mCapacity = newCap
    Debug.Print "INGEST: buffer grown to " & mCapacity
End Sub

Private Sub SetStage(ByVal stageName As String)
    mStage = stageName
    Debug.Print "INGEST STAGE: " & stageName
    On Error Resume Next
    mApp.SetStatus "Ingest: " & stageName
    On Error GoTo 0
End Sub

Private Sub ReportFailure(ByVal context As String)
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim msg As String

    ' Capture before Restore — Restore uses On Error Resume Next and clears Err.
    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source
    If Len(Trim$(errDesc)) = 0 Then errDesc = "(no description)"
    If Len(Trim$(errSrc)) = 0 Then errSrc = "(unknown)"

    Debug.Print "========== INGEST FAILURE =========="
    Debug.Print "Context: " & context
    Debug.Print "Stage:   " & mStage
    Debug.Print "Number:  " & CStr(errNum)
    Debug.Print "Source:  " & errSrc
    Debug.Print "Desc:    " & errDesc
    Debug.Print "Files:   " & CStr(mFileCount) & "  Folders: " & CStr(mFolderCount) & "  Entries: " & CStr(mCount)
    Debug.Print "===================================="

    On Error Resume Next
    If Not mMap Is Nothing Then mMap.ProtectWorkbookUi
    mApp.Restore
    On Error GoTo 0

    msg = "Ingestion failed." & vbCrLf & vbCrLf & _
          "Context: " & context & vbCrLf & _
          "Stage: " & mStage & vbCrLf & _
          "Error #" & CStr(errNum) & vbCrLf & _
          "Source: " & errSrc & vbCrLf & _
          "Description: " & errDesc & vbCrLf & vbCrLf & _
          "Progress before fail: " & CStr(mFileCount) & " files, " & CStr(mFolderCount) & " folders." & vbCrLf & vbCrLf & _
          "Open the VBA editor and press Ctrl+G for the Immediate Window log."

    MsgBox msg, vbCritical, "LAN Search - Ingestion"
End Sub

Private Sub LogSkip(ByVal path As String, ByVal reason As String)
    If mSkipLog Is Nothing Then Set mSkipLog = CreateObject("Scripting.Dictionary")
    If Not mSkipLog.Exists(path) Then mSkipLog.Add path, reason
End Sub

Private Sub DumpSkipLog()
    Dim k As Variant
    If mSkipLog Is Nothing Then Exit Sub
    For Each k In mSkipLog.Keys
        Debug.Print "SKIP: " & CStr(k) & " :: " & CStr(mSkipLog(k))
    Next k
End Sub

Private Sub ClearSheetFilters(ByVal ws As Worksheet)
    Dim lo As ListObject
    On Error Resume Next
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    If ws.FilterMode Then ws.ShowAllData
    For Each lo In ws.ListObjects
        If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
        ' Drop AutoFilter on the table so bulk writes are allowed
        lo.ShowAutoFilter = False
        lo.ShowAutoFilter = True
    Next lo
    On Error GoTo 0
End Sub

Private Sub ApplyIndexWrite(ByVal uncRoot As String, ByVal mode As String)
    Dim ws As Worksheet
    Dim keep() As Variant
    Dim keepCount As Long
    Dim combined() As Variant
    Dim total As Long
    Dim i As Long
    Dim dest As Range

    Set ws = mMap.Database
    ClearSheetFilters ws

    ' Collect rows to preserve while the old table still exists.
    SetStage "Write Database: read keep rows (" & mode & ")"
    If mode = "Append" Then
        keepCount = ReadAllExistingFiles(ws, keep)
    Else
        keepCount = ReadFilesOutsideRoot(ws, uncRoot, keep)
    End If
    Debug.Print "INGEST: keepRows=" & CStr(keepCount) & " newRows=" & CStr(mCount)

    ' Drop the ListObject entirely before bulk write — DataBodyRange.Delete on
    ' large tables (and writing into a shrunk table) commonly raises 1004.
    SetStage "Write Database: wipe tblFiles"
    WipeFilesBody ws
    EnsureFilesHeader ws

    total = keepCount + mCount
    If total = 0 Then
        RebuildFilesTable ws, 0
        modIndexCache.InvalidateIndex
        Exit Sub
    End If

    SetStage "Write Database: build combined array total=" & CStr(total)
    ReDim combined(1 To total, 1 To COLS_FILES)
    For i = 1 To keepCount
        combined(i, 1) = keep(i, 1)
        combined(i, 2) = keep(i, 2)
        combined(i, 3) = modPathUtil.CoerceStoredSizeMb(keep(i, 3))
        If IsEmpty(keep(i, 4)) Or Len(Trim$(CStr(keep(i, 4) & ""))) = 0 Then
            combined(i, 4) = ENTRY_FILE
        Else
            combined(i, 4) = UCase$(Trim$(CStr(keep(i, 4))))
        End If
    Next i
    For i = 1 To mCount
        combined(keepCount + i, 1) = mBuffer(i, 1)
        combined(keepCount + i, 2) = mBuffer(i, 2)
        combined(keepCount + i, 3) = mBuffer(i, 3)
        combined(keepCount + i, 4) = mBuffer(i, 4)
    Next i

    SetStage "Write Database: dump " & CStr(total) & " rows to sheet"
    Set dest = ws.Range(ws.Cells(2, 1), ws.Cells(1 + total, COLS_FILES))
    dest.Value = combined
    ws.Range(ws.Cells(2, 3), ws.Cells(1 + total, 3)).NumberFormat = "#,##0.00"

    SetStage "Write Database: rebuild ListObject"
    RebuildFilesTable ws, total

    SetStage "Write Database: seed index cache"
    modIndexCache.SetIndexData combined, total
End Sub

Private Function ReadAllExistingFiles(ByVal ws As Worksheet, ByRef keep() As Variant) As Long
    Dim lastRow As Long
    Dim r As Long
    Dim n As Long
    Dim path As String
    Dim lastCol As Long

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        ReadAllExistingFiles = 0
        Exit Function
    End If

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 3 Then lastCol = 3

    ReDim keep(1 To lastRow - 1, 1 To COLS_FILES)
    n = 0
    For r = 2 To lastRow
        path = Trim$(CStr(ws.Cells(r, 1).Value & ""))
        If Len(path) = 0 Then GoTo NextAll
        n = n + 1
        keep(n, 1) = ws.Cells(r, 1).Value
        keep(n, 2) = ws.Cells(r, 2).Value
        keep(n, 3) = ws.Cells(r, 3).Value
        If lastCol >= 4 Then
            keep(n, 4) = ws.Cells(r, 4).Value
        Else
            keep(n, 4) = ENTRY_FILE
        End If
NextAll:
    Next r
    ReadAllExistingFiles = n
End Function

Private Function ReadFilesOutsideRoot(ByVal ws As Worksheet, ByVal uncRoot As String, ByRef keep() As Variant) As Long
    Dim lastRow As Long
    Dim r As Long
    Dim n As Long
    Dim path As String
    Dim lastCol As Long

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        ReadFilesOutsideRoot = 0
        Exit Function
    End If

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 3 Then lastCol = 3

    ReDim keep(1 To lastRow - 1, 1 To COLS_FILES)
    n = 0
    For r = 2 To lastRow
        path = Trim$(CStr(ws.Cells(r, 1).Value & ""))
        If Len(path) = 0 Then GoTo NextKeep
        If modPathUtil.PathStartsWithRoot(path, uncRoot) Then GoTo NextKeep
        n = n + 1
        keep(n, 1) = ws.Cells(r, 1).Value
        keep(n, 2) = ws.Cells(r, 2).Value
        keep(n, 3) = ws.Cells(r, 3).Value
        If lastCol >= 4 Then
            keep(n, 4) = ws.Cells(r, 4).Value
        Else
            keep(n, 4) = ENTRY_FILE
        End If
NextKeep:
    Next r
    ReadFilesOutsideRoot = n
End Function

Private Sub WipeFilesBody(ByVal ws As Worksheet)
    Dim tbl As ListObject
    Dim lastRow As Long
    Dim lastCol As Long

    ClearSheetFilters ws

    ' Delete the whole ListObject (not DataBodyRange.Delete) — large body deletes
    ' and half-resized tables are a frequent source of error 1004.
    On Error Resume Next
    Set tbl = ws.ListObjects(mMap.TableFilesName)
    If Not tbl Is Nothing Then
        tbl.ShowAutoFilter = False
        tbl.Delete
    End If
    On Error GoTo 0

    lastCol = COLS_FILES
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        ' Also clear other columns in case path column was empty but others weren't
        On Error Resume Next
        lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
        On Error GoTo 0
    End If
    If lastRow >= 2 Then
        ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).ClearContents
    End If
End Sub

Private Sub RebuildFilesTable(ByVal ws As Worksheet, ByVal dataRows As Long)
    Dim tbl As ListObject
    Dim lastRow As Long

    EnsureFilesHeader ws
    lastRow = Application.Max(2, 1 + dataRows)

    ' Always recreate cleanly after a plain-range dump.
    On Error Resume Next
    Set tbl = ws.ListObjects(mMap.TableFilesName)
    If Not tbl Is Nothing Then tbl.Delete
    Set tbl = Nothing
    On Error GoTo 0

    Set tbl = ws.ListObjects.Add(xlSrcRange, _
        ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, COLS_FILES)), , xlYes)
    tbl.Name = mMap.TableFilesName
End Sub

Private Sub EnsureFilesHeader(ByVal ws As Worksheet)
    ws.Cells(1, 1).Value = "FilePath"
    ws.Cells(1, 2).Value = "FileDate"
    ws.Cells(1, 3).Value = "SizeMB"
    ws.Cells(1, 4).Value = "EntryType"
    ws.Columns(3).NumberFormat = "#,##0.00"
End Sub

Private Sub EnsureFilesTableEmpty()
    Dim ws As Worksheet
    Set ws = mMap.Database
    WipeFilesBody ws
    EnsureFilesHeader ws
    RebuildFilesTable ws, 0
    modIndexCache.InvalidateIndex
End Sub

Private Function GetFilesTable() As ListObject
    On Error Resume Next
    Set GetFilesTable = mMap.Database.ListObjects(mMap.TableFilesName)
    On Error GoTo 0
End Function

Private Sub UpsertIngestedList(ByVal uncRoot As String, ByVal whenIngested As Date)
    Dim tbl As ListObject
    Dim storeRoot As String
    Dim r As Long
    Dim foundRow As Long
    Dim lr As ListRow
    Dim existing As String
    Dim existingUnc As String
    Dim totalMb As Double

    storeRoot = uncRoot
    If Right$(storeRoot, 1) = "\" Then storeRoot = Left$(storeRoot, Len(storeRoot) - 1)
    totalMb = Round(mTotalBytes / BYTES_PER_MB, 2)

    ClearSheetFilters mMap.Ingestion
    Set tbl = GetOrCreateIngestedTable()

    foundRow = 0
    If Not tbl.DataBodyRange Is Nothing Then
        For r = 1 To tbl.DataBodyRange.Rows.Count
            existing = CStr(tbl.DataBodyRange.Cells(r, 1).Value & "")
            existingUnc = modPathUtil.ToUncPath(existing, mMap.UncRootOverride)
            If Right$(existingUnc, 1) = "\" Then existingUnc = Left$(existingUnc, Len(existingUnc) - 1)
            If StrComp(existingUnc, storeRoot, vbTextCompare) = 0 _
               Or StrComp(existing, storeRoot, vbTextCompare) = 0 _
               Or StrComp(modPathUtil.ToDisplayPath(existingUnc), modPathUtil.ToDisplayPath(storeRoot), vbTextCompare) = 0 Then
                foundRow = r
                Exit For
            End If
        Next r
    End If

    If foundRow > 0 Then
        With tbl.DataBodyRange
            .Cells(foundRow, 1).Value = storeRoot
            .Cells(foundRow, 2).Value = whenIngested
            .Cells(foundRow, 3).Value = mFileCount
            .Cells(foundRow, 4).Value = mFolderCount
            .Cells(foundRow, 5).Value = totalMb
        End With
    Else
        Set lr = tbl.ListRows.Add
        With lr.Range
            .Cells(1, 1).Value = storeRoot
            .Cells(1, 2).Value = whenIngested
            .Cells(1, 3).Value = mFileCount
            .Cells(1, 4).Value = mFolderCount
            .Cells(1, 5).Value = totalMb
        End With
    End If

    ' Root ingest absorbs prior subfolder ingests under that root
    RemoveChildIngestedRoots tbl, storeRoot
    SortIngestedTable
    ApplyIngestedColumnFormats tbl
    RefreshIngestedSummary tbl
End Sub

Private Sub ApplyIngestedColumnFormats(ByVal tbl As ListObject)
    On Error Resume Next
    If Not tbl.DataBodyRange Is Nothing Then
        tbl.ListColumns(3).DataBodyRange.NumberFormat = "#,##0"      ' FileCount
        tbl.ListColumns(4).DataBodyRange.NumberFormat = "#,##0"      ' FolderCount
        tbl.ListColumns(5).DataBodyRange.NumberFormat = "#,##0.00"   ' TotalSizeMB
    End If
    On Error GoTo 0
End Sub

' Totals in E2:E5 only. Row 1 header + D2:D5 labels and all formatting are user-owned.
Private Sub RefreshIngestedSummary(Optional ByVal tbl As ListObject = Nothing)
    Dim ws As Worksheet
    Dim r As Long
    Dim scans As Long
    Dim files As Double
    Dim folders As Double
    Dim sizeMb As Double

    Set ws = mMap.Ingestion
    If tbl Is Nothing Then Set tbl = GetOrCreateIngestedTable()

    scans = 0
    files = 0
    folders = 0
    sizeMb = 0

    If Not tbl Is Nothing Then
        If Not tbl.DataBodyRange Is Nothing Then
            scans = tbl.DataBodyRange.Rows.Count
            For r = 1 To scans
                files = files + CDbl(tbl.DataBodyRange.Cells(r, 3).Value & 0)
                folders = folders + CDbl(tbl.DataBodyRange.Cells(r, 4).Value & 0)
                sizeMb = sizeMb + CDbl(tbl.DataBodyRange.Cells(r, 5).Value & 0)
            Next r
        End If
    End If

    ws.Range(mMap.IngestSummaryScansValue).Value = scans
    ws.Range(mMap.IngestSummaryFilesValue).Value = files
    ws.Range(mMap.IngestSummaryFoldersValue).Value = folders
    ws.Range(mMap.IngestSummarySizeValue).Value = Round(sizeMb, 2)
End Sub

Private Sub RemoveChildIngestedRoots(ByVal tbl As ListObject, ByVal parentRoot As String)
    Dim r As Long
    Dim existing As String
    Dim existingNorm As String
    Dim parentNorm As String

    If tbl.DataBodyRange Is Nothing Then Exit Sub

    parentNorm = NormalizeRootKey(parentRoot)
    For r = tbl.DataBodyRange.Rows.Count To 1 Step -1
        existing = Trim$(CStr(tbl.DataBodyRange.Cells(r, 1).Value & ""))
        If Len(existing) = 0 Then
            tbl.ListRows(r).Delete
            GoTo NextChild
        End If
        existingNorm = NormalizeRootKey(existing)
        If StrComp(existingNorm, parentNorm, vbTextCompare) = 0 Then GoTo NextChild
        If modPathUtil.PathStartsWithRoot(existingNorm, parentNorm) Then
            Debug.Print "INGEST: removing child root from tblIngested: " & existing
            tbl.ListRows(r).Delete
        End If
NextChild:
    Next r
End Sub

Private Function NormalizeRootKey(ByVal rootPath As String) As String
    Dim unc As String
    unc = modPathUtil.ToUncPath(rootPath, mMap.UncRootOverride)
    unc = Replace(unc, "/", "\")
    Do While Right$(unc, 1) = "\"
        unc = Left$(unc, Len(unc) - 1)
    Loop
    NormalizeRootKey = unc
End Function

Private Function GetOrCreateIngestedTable() As ListObject
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim lo As ListObject
    Dim headerRow As Long
    Dim startCol As Long
    Dim colCount As Long

    Set ws = mMap.Ingestion
    headerRow = mMap.IngestedHeaderRow
    startCol = mMap.IngestedStartCol
    colCount = mMap.IngestedColCount

    On Error Resume Next
    Set tbl = ws.ListObjects(mMap.TableIngestedName)
    On Error GoTo 0

    If tbl Is Nothing Then
        For Each lo In ws.ListObjects
            If StrComp(CStr(lo.HeaderRowRange.Cells(1, 1).Value & ""), "RootPath", vbTextCompare) = 0 Then
                Set tbl = lo
                On Error Resume Next
                tbl.Name = mMap.TableIngestedName
                On Error GoTo 0
                Exit For
            End If
        Next lo
    End If

    ' Move existing table to A6 if it lives elsewhere
    If Not tbl Is Nothing Then
        If tbl.HeaderRowRange.Row <> headerRow Or tbl.HeaderRowRange.Column <> startCol Then
            Set tbl = RelocateIngestedTable(tbl, headerRow, startCol, colCount)
        End If
    End If

    If tbl Is Nothing Then
        WriteIngestedHeaders ws, headerRow, startCol
        Set tbl = ws.ListObjects.Add(xlSrcRange, _
            ws.Range(ws.Cells(headerRow, startCol), ws.Cells(headerRow + 1, startCol + colCount - 1)), , xlYes)
        On Error Resume Next
        tbl.Name = mMap.TableIngestedName
        On Error GoTo 0
        If Not tbl.DataBodyRange Is Nothing Then
            If WorksheetFunction.CountA(tbl.DataBodyRange) = 0 Then
                On Error Resume Next
                tbl.DataBodyRange.Rows(1).Delete
                On Error GoTo 0
            End If
        End If
    End If

    If tbl.ListColumns.Count < colCount Then
        Err.Raise vbObjectError + 3, "clsFileIndexer", _
            "Ingested table '" & tbl.Name & "' has " & tbl.ListColumns.Count & _
            " columns; need " & colCount & " (RootPath, DateIngested, FileCount, FolderCount, TotalSizeMB)."
    End If

    Set GetOrCreateIngestedTable = tbl
End Function

Private Sub WriteIngestedHeaders(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal startCol As Long)
    ws.Cells(headerRow, startCol).Value = "RootPath"
    ws.Cells(headerRow, startCol + 1).Value = "DateIngested"
    ws.Cells(headerRow, startCol + 2).Value = "FileCount"
    ws.Cells(headerRow, startCol + 3).Value = "FolderCount"
    ws.Cells(headerRow, startCol + 4).Value = "TotalSizeMB"
End Sub

Private Function RelocateIngestedTable(ByVal tbl As ListObject, ByVal headerRow As Long, _
                                      ByVal startCol As Long, ByVal colCount As Long) As ListObject
    Dim ws As Worksheet
    Dim oldRange As Range
    Dim data As Variant
    Dim rowCount As Long
    Dim newTbl As ListObject

    Set ws = tbl.Parent
    Set oldRange = tbl.Range
    rowCount = oldRange.Rows.Count
    data = oldRange.Value

    On Error Resume Next
    tbl.Unlist
    On Error GoTo 0
    oldRange.Clear

    If IsArray(data) Then
        ws.Range(ws.Cells(headerRow, startCol), _
                 ws.Cells(headerRow + rowCount - 1, startCol + colCount - 1)).Value = data
    Else
        WriteIngestedHeaders ws, headerRow, startCol
    End If

    Set newTbl = ws.ListObjects.Add(xlSrcRange, _
        ws.Range(ws.Cells(headerRow, startCol), _
                 ws.Cells(headerRow + Application.Max(rowCount, 2) - 1, startCol + colCount - 1)), , xlYes)
    On Error Resume Next
    newTbl.Name = mMap.TableIngestedName
    On Error GoTo 0

    Set RelocateIngestedTable = newTbl
End Function

Private Sub SortIngestedTable()
    Dim tbl As ListObject
    Set tbl = GetOrCreateIngestedTable()
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    With tbl.Sort
        .SortFields.Clear
        .SortFields.Add Key:=tbl.ListColumns(1).Range, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Header = xlYes
        .Apply
    End With
End Sub
