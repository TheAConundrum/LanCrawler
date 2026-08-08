Attribute VB_Name = "modIndexCache"
Option Explicit
Option Private Module

' Workbook-lifetime cache of Database!tblFiles for search.
' Loaded on first search; seeded/replaced after a successful ingest write;
' cleared when the index is emptied. Lives until the workbook closes.
'
' Parallel search arrays (path/name/flex/folder/size/date) are built once on load
' so each search avoids re-parsing rows and re-normalizing names.

Private Const SHEET_DATABASE As String = "Database"
Private Const TABLE_FILES As String = "tblFiles"

Private mData As Variant
Private mRowCount As Long
Private mLoaded As Boolean

' Parallel search arrays (1-based), built with mData
Private mPath() As String
Private mName() As String
Private mNameFlex() As String
Private mIsFolder() As Boolean
Private mSizeMb() As Double
Private mDate() As Variant
Private mParentIdx() As Long          ' immediate parent FOLDER row index (0 = none)
Private mDescFileCount() As Long      ' indexed files under folder (recursive)
Private mDescFolderCount() As Long    ' indexed subfolders under folder (recursive)
Private mDescIndexedSize() As Double  ' sum of indexed file SizeMB under folder (legacy fallback)
Private mFolderRowByPath As Object    ' Scripting.Dictionary: folder UNC -> row index
Private mShareRoot() As String        ' UNC share root per row (\\server\share)
Private mDriveByLabel As Object       ' Scripting.Dictionary: friendly label -> share UNC
Private mLabelByShare As Object       ' Scripting.Dictionary: share UNC -> friendly label
Private mExtrasReady As Boolean

' Last GetIndexData / LoadFromDatabase diagnostics (seconds)
Private mLastGetWasHit As Boolean
Private mLastGetSeconds As Double
Private mLastLoadSeconds As Double

Public Function IsIndexLoaded() As Boolean
    IsIndexLoaded = mLoaded
End Function

Public Function IndexRowCount() As Long
    If Not mLoaded Then
        IndexRowCount = 0
    Else
        IndexRowCount = mRowCount
    End If
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

' Ensure cache is loaded (no array copy). Prefer this for search.
Public Sub EnsureLoaded(Optional ByVal wb As Workbook = Nothing)
    Dim t0 As Double
    t0 = Timer
    mLastGetWasHit = mLoaded
    If Not mLoaded Then
        LoadFromDatabase wb
        mLastGetSeconds = ElapsedSeconds(t0)
        Debug.Print "INDEX CACHE: EnsureLoaded MISS rows=" & CStr(mRowCount) & _
                    " " & FormatSeconds(mLastGetSeconds)
    Else
        mLastGetSeconds = ElapsedSeconds(t0)
        Debug.Print "INDEX CACHE: EnsureLoaded HIT rows=" & CStr(mRowCount) & _
                    " " & FormatSeconds(mLastGetSeconds)
        ' Keep drive dropdown in sync even on cache hits
        RefreshDashboardDriveList wb
    End If
End Sub

' Returns cached rows (1-based 2D). Loads from Database sheet on first use.
' Note: assigning the Variant return value copies the array — prefer EnsureLoaded + CollectNameHits.
Public Function GetIndexData(ByRef rowCount As Long) As Variant
    Dim t0 As Double
    Dim wasLoaded As Boolean

    t0 = Timer
    wasLoaded = mLoaded
    mLastGetWasHit = mLoaded

    If Not mLoaded Then LoadFromDatabase ActiveWorkbook

    rowCount = mRowCount
    If mRowCount = 0 Then
        GetIndexData = Empty
    Else
        GetIndexData = mData
    End If

    mLastGetSeconds = ElapsedSeconds(t0)
    If wasLoaded Then
        Debug.Print "INDEX CACHE: GetIndexData HIT rows=" & CStr(mRowCount) & _
                    " copy+return=" & FormatSeconds(mLastGetSeconds)
    Else
        Debug.Print "INDEX CACHE: GetIndexData MISS (loaded) rows=" & CStr(mRowCount) & _
                    " load=" & FormatSeconds(mLastLoadSeconds) & _
                    " getTotal=" & FormatSeconds(mLastGetSeconds)
    End If
End Function

' Replace cache with a full index array (e.g. after ingest write).
Public Sub SetIndexData(ByRef source As Variant, ByVal rowCount As Long)
    Dim t0 As Double
    If rowCount <= 0 Or Not IsArray(source) Then
        InvalidateIndex
        Exit Sub
    End If
    mData = source
    mRowCount = rowCount
    mLoaded = True
    mLastLoadSeconds = 0
    t0 = Timer
    BuildSearchExtras
    Debug.Print "INDEX CACHE: set from ingest/array rows=" & CStr(mRowCount) & _
                " extras=" & FormatSeconds(ElapsedSeconds(t0))
End Sub

Public Sub InvalidateIndex()
    On Error Resume Next
    Erase mData
    Erase mPath
    Erase mName
    Erase mNameFlex
    Erase mIsFolder
    Erase mSizeMb
    Erase mDate
    Erase mParentIdx
    Erase mDescFileCount
    Erase mDescFolderCount
    Erase mDescIndexedSize
    Erase mShareRoot
    Set mFolderRowByPath = Nothing
    Set mDriveByLabel = Nothing
    Set mLabelByShare = Nothing
    On Error GoTo 0
    mData = Empty
    mRowCount = 0
    mLoaded = False
    mExtrasReady = False
    Debug.Print "INDEX CACHE: invalidated"
    ' Do not refresh drive dropdown here — BuildSearchExtras does it after drives are known.
End Sub

Public Sub LoadFromDatabase(Optional ByVal wb As Workbook = Nothing)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim data As Variant
    Dim n As Long
    Dim cols As Long
    Dim r As Long
    Dim normalized() As Variant
    Dim t0 As Double
    Dim tRead As Double
    Dim tNorm As Double
    Dim tExtras As Double
    Dim secRead As Double
    Dim secNorm As Double
    Dim secExtras As Double

    t0 = Timer
    InvalidateIndex

    If wb Is Nothing Then Set wb = ActiveWorkbook
    If wb Is Nothing Then
        mLastLoadSeconds = ElapsedSeconds(t0)
        Debug.Print "INDEX CACHE: LoadFromDatabase aborted (no workbook) " & FormatSeconds(mLastLoadSeconds)
        Exit Sub
    End If

    On Error Resume Next
    Set ws = wb.Worksheets(SHEET_DATABASE)
    If ws Is Nothing Then
        mLastLoadSeconds = ElapsedSeconds(t0)
        Debug.Print "INDEX CACHE: LoadFromDatabase aborted (no Database sheet) " & FormatSeconds(mLastLoadSeconds)
        Exit Sub
    End If
    Set tbl = ws.ListObjects(TABLE_FILES)
    On Error GoTo 0

    If tbl Is Nothing Then
        mLastLoadSeconds = ElapsedSeconds(t0)
        Debug.Print "INDEX CACHE: LoadFromDatabase aborted (no tblFiles) " & FormatSeconds(mLastLoadSeconds)
        Exit Sub
    End If
    If tbl.DataBodyRange Is Nothing Then
        mLastLoadSeconds = ElapsedSeconds(t0)
        Debug.Print "INDEX CACHE: LoadFromDatabase empty table " & FormatSeconds(mLastLoadSeconds)
        Exit Sub
    End If
    If tbl.ListRows.Count = 0 Then
        mLastLoadSeconds = ElapsedSeconds(t0)
        Debug.Print "INDEX CACHE: LoadFromDatabase 0 rows " & FormatSeconds(mLastLoadSeconds)
        Exit Sub
    End If

    On Error Resume Next
    Application.StatusBar = "LAN Search: loading index cache..."
    On Error GoTo 0

    tRead = Timer
    data = tbl.DataBodyRange.Value
    secRead = ElapsedSeconds(tRead)

    If Not IsArray(data) Then
        ClearAppStatusBar
        mLastLoadSeconds = ElapsedSeconds(t0)
        Debug.Print "INDEX CACHE: LoadFromDatabase non-array body " & FormatSeconds(mLastLoadSeconds)
        Exit Sub
    End If

    n = UBound(data, 1)
    cols = UBound(data, 2)

    tNorm = Timer
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
        If r Mod 2500 = 0 Then DoEvents
    Next r
    secNorm = ElapsedSeconds(tNorm)

    mData = normalized
    mRowCount = n
    mLoaded = True

    tExtras = Timer
    BuildSearchExtras
    secExtras = ElapsedSeconds(tExtras)

    mLastLoadSeconds = ElapsedSeconds(t0)

    Debug.Print "INDEX CACHE: LoadFromDatabase rows=" & CStr(mRowCount) & _
                " sheetRead=" & FormatSeconds(secRead) & _
                " normalize=" & FormatSeconds(secNorm) & _
                " searchExtras=" & FormatSeconds(secExtras) & _
                " total=" & FormatSeconds(mLastLoadSeconds)

    ClearAppStatusBar
End Sub

Private Sub ClearAppStatusBar()
    On Error Resume Next
    Application.DisplayStatusBar = True
    Application.StatusBar = vbNullString
    Application.StatusBar = False
    On Error GoTo 0
End Sub

' No-op if already warm (used by deferred Workbook_Open load).
Public Sub WarmIndexIfNeeded(Optional ByVal wb As Workbook = Nothing)
    If mLoaded Then
        Debug.Print "INDEX CACHE: already warm rows=" & CStr(mRowCount)
        Exit Sub
    End If
    Debug.Print "INDEX CACHE: WarmIndexIfNeeded starting load..."
    LoadFromDatabase wb
End Sub

' Fast name scan against precomputed arrays (no Variant array copy).
' shareFilter: UNC share root to restrict (empty = All drives).
Public Sub CollectNameHits(ByVal term1 As String, ByVal op As String, ByVal term2 As String, _
                           ByVal flexible As Boolean, _
                           ByVal includeFiles As Boolean, ByVal includeFolders As Boolean, _
                           ByVal hasSizeFilter As Boolean, ByVal sizeOp As String, ByVal sizeMb As Double, _
                           ByVal hits As Object, ByVal folderHits As Object, _
                           Optional ByVal shareFilter As String = "")
    Dim i As Long
    Dim term1N As String
    Dim term2N As String
    Dim hasTerm2 As Boolean
    Dim matched As Boolean
    Dim has1 As Boolean
    Dim has2 As Boolean
    Dim hay As String
    Dim filterShare As String

    If Not mLoaded Or mRowCount = 0 Then Exit Sub
    If Not mExtrasReady Then BuildSearchExtras

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
        If Len(mPath(i)) = 0 Then GoTo NextRow

        If Len(filterShare) > 0 Then
            If StrComp(mShareRoot(i), filterShare, vbTextCompare) <> 0 Then GoTo NextRow
        End If

        If mIsFolder(i) Then
            If Not includeFolders Then GoTo NextRow
        Else
            If Not includeFiles Then GoTo NextRow
        End If

        If flexible Then
            hay = mNameFlex(i)
            has1 = (InStr(1, hay, term1N, vbBinaryCompare) > 0)
            If hasTerm2 Then
                has2 = (InStr(1, hay, term2N, vbBinaryCompare) > 0)
                matched = CombineMatch(has1, has2, op)
            Else
                matched = has1
            End If
        Else
            hay = mName(i)
            has1 = (InStr(1, hay, term1, vbTextCompare) > 0)
            If hasTerm2 Then
                has2 = (InStr(1, hay, term2, vbTextCompare) > 0)
                matched = CombineMatch(has1, has2, op)
            Else
                matched = has1
            End If
        End If

        If Not matched Then GoTo NextRow

        If mIsFolder(i) Then
            If Not folderHits.Exists(mPath(i)) Then
                folderHits.Add mPath(i), Array(mDate(i), mSizeMb(i))
            End If
        Else
            If hasSizeFilter Then
                If Not SizeMatchesLocal(mSizeMb(i), sizeOp, sizeMb) Then GoTo NextRow
            End If
            If Not hits.Exists(mPath(i)) Then
                hits.Add mPath(i), Array(mDate(i), mSizeMb(i), False, 0&, 0&)
            End If
        End If
NextRow:
    Next i
End Sub

' Map Dashboard drive dropdown label to UNC share root (empty = All).
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
        ' Allow pasting a UNC share / path directly
        ResolveDriveShareFilter = modPathUtil.UncShareRoot(v)
    End If
End Function

' Friendly drive label for results (e.g. SHARE_Public). Uses cache built on warm load.
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

' Rebuild drive dropdown on Dashboard H5 from unique shares in the loaded index.
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
    Dim share As String
    Dim label As String
    Dim unique As String
    Dim suffix As Long

    On Error GoTo FailQuiet
    If wb Is Nothing Then Set wb = ActiveWorkbook
    If wb Is Nothing Then
        Debug.Print "INDEX CACHE: RefreshDashboardDriveList aborted (no workbook)"
        Exit Sub
    End If

    ' Recover drive map if somehow empty while index rows exist
    If mLoaded And mRowCount > 0 Then
        If mDriveByLabel Is Nothing Then
            Set mDriveByLabel = CreateObject("Scripting.Dictionary")
            mDriveByLabel.CompareMode = 1
        End If
        If mLabelByShare Is Nothing Then
            Set mLabelByShare = CreateObject("Scripting.Dictionary")
            mLabelByShare.CompareMode = 1
        End If
        If mDriveByLabel.Count = 0 Then
            If ShareRootArrayReady() Then
                On Error Resume Next
                For i = 1 To mRowCount
                    share = mShareRoot(i)
                    If Len(share) > 0 Then
                        If Not mLabelByShare.Exists(share) Then
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
                        End If
                    End If
                Next i
                On Error GoTo FailQuiet
                Debug.Print "INDEX CACHE: rebuilt drive map count=" & CStr(mDriveByLabel.Count)
            End If
        End If
    End If

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
    If Err.Number <> 0 Then
        Debug.Print "INDEX CACHE: drive Validation.Add failed Err=" & CStr(Err.Number) & _
                    " " & Err.Description & " formula=" & Left$(formula, 120)
        On Error GoTo FailQuiet
        Err.Raise Err.Number, , Err.Description
    End If
    On Error GoTo FailQuiet

    With cell.Validation
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = False
        .ShowError = False
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
        activeWas.Activate
        On Error GoTo FailQuiet
    End If

    Debug.Print "INDEX CACHE: drive dropdown @H5 items=" & CStr(n + 1) & _
                " formula=" & formula
    Exit Sub
FailQuiet:
    Debug.Print "INDEX CACHE: RefreshDashboardDriveList skipped Err=" & CStr(Err.Number) & _
                " " & Err.Description
End Sub

Private Function ShareRootArrayReady() As Boolean
    On Error Resume Next
    ShareRootArrayReady = (UBound(mShareRoot) >= 1)
    On Error GoTo 0
End Function

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

' O(folderHits) using counts precomputed in BuildSearchExtras.
Public Sub CommitFolderHitsFast(ByVal hits As Object, ByVal folderHits As Object, _
                                ByVal hasSizeFilter As Boolean, ByVal sizeOp As String, ByVal sizeMb As Double)
    Dim key As Variant
    Dim folderUnc As String
    Dim meta As Variant
    Dim folderDate As Variant
    Dim storedMb As Double
    Dim displayMb As Double
    Dim rowIdx As Long
    Dim fc As Long
    Dim dc As Long
    Dim lo As Long
    Dim hi As Long

    If folderHits Is Nothing Then Exit Sub
    If folderHits.Count = 0 Then Exit Sub
    If Not mLoaded Or mRowCount = 0 Then Exit Sub
    If Not mExtrasReady Then BuildSearchExtras

    For Each key In folderHits.Keys
        folderUnc = CStr(key)
        If hits.Exists(folderUnc) Then GoTo NextFolder

        meta = folderHits(key)
        folderDate = Empty
        storedMb = 0.01
        If IsArray(meta) Then
            lo = LBound(meta)
            hi = UBound(meta)
            If hi >= lo Then
                If Not IsEmpty(meta(lo)) Then folderDate = meta(lo)
            End If
            If hi >= lo + 1 Then
                storedMb = modPathUtil.CoerceStoredSizeMb(meta(lo + 1))
            End If
        End If

        rowIdx = 0
        If Not mFolderRowByPath Is Nothing Then
            If mFolderRowByPath.Exists(folderUnc) Then rowIdx = CLng(mFolderRowByPath(folderUnc))
        End If
        If rowIdx < 1 Or rowIdx > mRowCount Then rowIdx = 0

        If storedMb > 0.01 Then
            displayMb = storedMb
        ElseIf rowIdx > 0 Then
            If mDescIndexedSize(rowIdx) > 0.01 Then
                displayMb = Round(mDescIndexedSize(rowIdx), 2)
            Else
                displayMb = 0.01
            End If
        Else
            displayMb = 0.01
        End If

        If hasSizeFilter Then
            If Not SizeMatchesLocal(displayMb, sizeOp, sizeMb) Then GoTo NextFolder
        End If

        If rowIdx > 0 Then
            fc = mDescFileCount(rowIdx)
            dc = mDescFolderCount(rowIdx)
        Else
            fc = 0
            dc = 0
        End If

        hits.Add folderUnc, Array(folderDate, displayMb, True, fc, dc)
NextFolder:
    Next key
End Sub

Private Sub BuildSearchExtras()
    Dim i As Long
    Dim j As Long
    Dim cols As Long
    Dim t0 As Double
    Dim tCounts As Double
    Dim p As String
    Dim nm As String
    Dim parent As String
    Dim secNames As Double
    Dim share As String
    Dim label As String
    Dim unique As String
    Dim suffix As Long
    Dim seenShares As Object

    mExtrasReady = False
    If Not mLoaded Or mRowCount <= 0 Or Not IsArray(mData) Then Exit Sub

    t0 = Timer
    cols = UBound(mData, 2)
    ReDim mPath(1 To mRowCount)
    ReDim mName(1 To mRowCount)
    ReDim mNameFlex(1 To mRowCount)
    ReDim mIsFolder(1 To mRowCount)
    ReDim mSizeMb(1 To mRowCount)
    ReDim mDate(1 To mRowCount)
    ReDim mParentIdx(1 To mRowCount)
    ReDim mDescFileCount(1 To mRowCount)
    ReDim mDescFolderCount(1 To mRowCount)
    ReDim mDescIndexedSize(1 To mRowCount)
    ReDim mShareRoot(1 To mRowCount)

    Set mFolderRowByPath = CreateObject("Scripting.Dictionary")
    mFolderRowByPath.CompareMode = 1
    Set mDriveByLabel = CreateObject("Scripting.Dictionary")
    mDriveByLabel.CompareMode = 1
    Set mLabelByShare = CreateObject("Scripting.Dictionary")
    mLabelByShare.CompareMode = 1
    Set seenShares = CreateObject("Scripting.Dictionary")
    seenShares.CompareMode = 1

    For i = 1 To mRowCount
        p = CStr(mData(i, 1) & "")
        If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
        mPath(i) = p
        nm = modPathUtil.FileNameOnly(p)
        mName(i) = nm
        mNameFlex(i) = modPathUtil.NormalizeForMatch(nm)
        If cols >= 4 Then
            mIsFolder(i) = (UCase$(Trim$(CStr(mData(i, 4) & ""))) = "FOLDER")
        Else
            mIsFolder(i) = False
        End If
        If cols >= 3 Then
            mSizeMb(i) = modPathUtil.CoerceStoredSizeMb(mData(i, 3))
        Else
            mSizeMb(i) = 0.01
        End If
        If cols >= 2 Then
            mDate(i) = mData(i, 2)
        Else
            mDate(i) = Empty
        End If
        mParentIdx(i) = 0
        mDescFileCount(i) = 0
        mDescFolderCount(i) = 0
        mDescIndexedSize(i) = 0
        share = modPathUtil.UncShareRoot(p)
        mShareRoot(i) = share
        If Len(share) > 0 Then
            If Not seenShares.Exists(share) Then
                seenShares.Add share, True
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
                Debug.Print "INDEX CACHE: drive map " & unique & " <= " & share
            End If
        End If
        If mIsFolder(i) And Len(p) > 0 Then
            If Not mFolderRowByPath.Exists(p) Then mFolderRowByPath.Add p, i
        End If
        If i Mod 2500 = 0 Then DoEvents
    Next i

    ' Immediate parent folder row for each entry
    For i = 1 To mRowCount
        parent = ParentNoSlash(mPath(i))
        If Len(parent) > 0 Then
            If mFolderRowByPath.Exists(parent) Then mParentIdx(i) = CLng(mFolderRowByPath(parent))
        End If
    Next i

    secNames = ElapsedSeconds(t0)

    ' Bubble descendant counts / indexed size up the parent chain (once per warm load)
    tCounts = Timer
    For i = 1 To mRowCount
        j = mParentIdx(i)
        Do While j > 0
            If mIsFolder(i) Then
                mDescFolderCount(j) = mDescFolderCount(j) + 1
            Else
                mDescFileCount(j) = mDescFileCount(j) + 1
                mDescIndexedSize(j) = mDescIndexedSize(j) + mSizeMb(i)
            End If
            j = mParentIdx(j)
        Loop
    Next i

    mExtrasReady = True
    Debug.Print "INDEX CACHE: BuildSearchExtras rows=" & CStr(mRowCount) & _
                " drives=" & CStr(mDriveByLabel.Count) & _
                " names=" & FormatSeconds(secNames) & _
                " descCounts=" & FormatSeconds(ElapsedSeconds(tCounts)) & _
                " total=" & FormatSeconds(ElapsedSeconds(t0))

    RefreshDashboardDriveList ActiveWorkbook
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

Private Function ParentNoSlash(ByVal fullPath As String) As String
    Dim parent As String
    parent = modPathUtil.ParentFolderPath(fullPath)
    If Right$(parent, 1) = "\" Then parent = Left$(parent, Len(parent) - 1)
    ParentNoSlash = parent
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
