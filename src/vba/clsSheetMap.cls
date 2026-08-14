VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSheetMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Resolves workbook sheet / table / cell layout for LAN Search Tool.
' Sheet names: Dashboard, Ingestion, Database

Private Const SHEET_DASHBOARD As String = "Dashboard"
Private Const SHEET_INGESTION As String = "Ingestion"
Private Const SHEET_DATABASE As String = "Database"

Private Const TABLE_FILES As String = "tblFiles"
Private Const TABLE_RESULTS As String = "tblResults"
Private Const TABLE_INGESTED As String = "tblIngested"

' Ingestion (folder path comes from Explorer picker on the ingest button)
' UncRoot / Mode sit in B (D:E reserved for compact summary).
Private Const ADDR_UNC_ROOT As String = "B2"
Private Const ADDR_INGEST_MODE As String = "B4" ' ReplaceRoot | Append | RebuildAll
Private Const INGESTED_HEADER_ROW As Long = 6 ' Links already Eaten headers start at A6
Private Const INGESTED_START_COL As Long = 1  ' A
Private Const INGESTED_COL_COUNT As Long = 5
' Compact summary: row 1 header (user-owned); D2:D5 labels, E2:E5 values (macro writes values only)
Private Const ADDR_SUM_SCANS_LBL As String = "D2"
Private Const ADDR_SUM_SCANS_VAL As String = "E2"
Private Const ADDR_SUM_FILES_LBL As String = "D3"
Private Const ADDR_SUM_FILES_VAL As String = "E3"
Private Const ADDR_SUM_FOLDERS_LBL As String = "D4"
Private Const ADDR_SUM_FOLDERS_VAL As String = "E4"
Private Const ADDR_SUM_SIZE_LBL As String = "D5"
Private Const ADDR_SUM_SIZE_VAL As String = "E5"

' Dashboard layout (user-owned chrome):
' Row 2 = criteria headers; row 3 = Contains A3:B3 | Op C3 | Contains D3:E3
'         Size is F3 (Over/Under) | MB G3 (number) | Flexible H3 | Search folders I3
' Row 4 = File Search / Drive labels; row 5 = File Search A5:G5 | Drive H5:I5
' Search stats values M3 / M4 / M5 (merged M:N); row 6 divider;
' Results headers row 7; data row 8+.
Private Const CRITERIA_HEADER_ROW As Long = 2
Private Const CRITERIA_FIRST_ROW As Long = 3
Private Const CRITERIA_LAST_ROW As Long = 3
Private Const CRITERIA_COL_TERM1 As Long = 1 ' A (merged A3:B3)
Private Const CRITERIA_COL_OP As Long = 3    ' C And/Or/Not
Private Const CRITERIA_COL_TERM2 As Long = 4 ' D (merged D3:E3)
Private Const CRITERIA_COL_SIZE_OP As Long = 6 ' F Size is (Over/Under)
Private Const CRITERIA_COL_SIZE_MB As Long = 7 ' G MB number
Private Const ADDR_TERM1 As String = "A3"
Private Const ADDR_OP As String = "C3"
Private Const ADDR_TERM2 As String = "D3"
Private Const ADDR_SIZE_OP As String = "F3"
Private Const ADDR_SIZE_MB As String = "G3"
Private Const ADDR_FLEXIBLE_MATCH As String = "H3" ' Yes/No
Private Const ADDR_SEARCH_FOLDERS As String = "I3" ' Yes/No/Only
Private Const ADDR_FILE_SEARCH As String = "A5"   ' merged A5:G5 (partial filename)
Private Const ADDR_DRIVE_FILTER As String = "H5"  ' merged H5:I5
' Search stats: values in merged M3:N3 / M4:N4 / M5:N5 (write top-left)
Private Const ADDR_STAT_FILES_VAL As String = "M3"
Private Const ADDR_STAT_FOLDERS_VAL As String = "M4"
Private Const ADDR_STAT_SIZE_VAL As String = "M5"
' Results headers row 7; grey divider row 6; data from row 8+.
' File A:K | Type L | Size MB M | Date N | Drive Address O | Folder Path P:T
Private Const RESULTS_HEADER_ROW As Long = 7
Private Const RESULTS_DIVIDER_ROW As Long = 6
Private Const RESULTS_FILE_COL As Long = 1           ' A
Private Const RESULTS_FILE_COL_END As Long = 11      ' K  (File merges A:K)
Private Const RESULTS_TYPE_COL As Long = 12          ' L File Type
Private Const RESULTS_SIZE_COL As Long = 13           ' M Size MB
Private Const RESULTS_DATE_COL As Long = 14          ' N Date Created
Private Const RESULTS_DRIVE_COL As Long = 15         ' O Drive Address (friendly label)
Private Const RESULTS_FOLDER_COL As Long = 16        ' P Folder Path
Private Const RESULTS_FOLDER_COL_END As Long = 20    ' T  (Folder Path merges P:T)
Private Const RESULTS_LAST_COL As Long = 20
' Hidden UNC targets for double-click open (not part of visible results band).
Private Const RESULTS_LINK_COL As Long = 21          ' U
' Soft cap for on-sheet results (merges/stripes are prebuilt once for this height).
Private Const RESULTS_MAX_ROWS As Long = 20000

Private mWorkbook As Workbook

Public Sub Init(Optional ByVal wb As Workbook = Nothing)
    If wb Is Nothing Then
        Set mWorkbook = ActiveWorkbook
    Else
        Set mWorkbook = wb
    End If
End Sub

Public Property Get Workbook() As Workbook
    Set Workbook = mWorkbook
End Property

Public Property Get Dashboard() As Worksheet
    Set Dashboard = mWorkbook.Worksheets(SHEET_DASHBOARD)
End Property

Public Property Get Ingestion() As Worksheet
    Set Ingestion = mWorkbook.Worksheets(SHEET_INGESTION)
End Property

Public Property Get Database() As Worksheet
    Set Database = mWorkbook.Worksheets(SHEET_DATABASE)
End Property

Public Property Get TableFilesName() As String
    TableFilesName = TABLE_FILES
End Property

Public Property Get TableResultsName() As String
    TableResultsName = TABLE_RESULTS
End Property

Public Property Get TableIngestedName() As String
    TableIngestedName = TABLE_INGESTED
End Property

Public Property Get UncRootOverride() As String
    UncRootOverride = Trim$(CStr(Ingestion.Range(ADDR_UNC_ROOT).Value & ""))
End Property

Public Property Get IngestMode() As String
    Dim v As String
    v = UCase$(Trim$(CStr(Ingestion.Range(ADDR_INGEST_MODE).Value & "")))
    Select Case v
        Case "APPEND"
            IngestMode = "Append"
        Case "REBUILDALL", "REBUILD ALL", "REBUILD"
            IngestMode = "RebuildAll"
        Case Else
            IngestMode = "ReplaceRoot"
    End Select
End Property

Public Property Get FlexibleMatch() As Boolean
    Dim v As String
    v = UCase$(Trim$(CStr(Dashboard.Range(ADDR_FLEXIBLE_MATCH).Value & "")))
    FlexibleMatch = (v = "YES" Or v = "Y" Or v = "TRUE" Or v = "1" Or v = "ON")
End Property

' Match folder names. Blank/Yes = files+folders; No = files only; Only = folders only.
Public Property Get SearchFoldersMode() As String
    Dim v As String
    v = UCase$(Trim$(CStr(Dashboard.Range(ADDR_SEARCH_FOLDERS).Value & "")))
    Select Case v
        Case "NO", "N", "FALSE", "0", "OFF"
            SearchFoldersMode = "No"
        Case "ONLY", "FOLDER", "FOLDERS"
            SearchFoldersMode = "Only"
        Case Else
            ' Yes / Y / True / blank (default)
            SearchFoldersMode = "Yes"
    End Select
End Property

' Drive filter dropdown value (friendly label). Blank or All = no filter.
Public Property Get DriveFilterLabel() As String
    Dim v As String
    v = Trim$(CStr(Dashboard.Range(ADDR_DRIVE_FILTER).Value & ""))
    If Len(v) = 0 Then
        DriveFilterLabel = "All"
    Else
        DriveFilterLabel = v
    End If
End Property

Public Property Get DriveFilterAddress() As String
    DriveFilterAddress = ADDR_DRIVE_FILTER
End Property

Public Property Get StatFilesValueAddress() As String
    StatFilesValueAddress = ADDR_STAT_FILES_VAL
End Property

Public Property Get StatFoldersValueAddress() As String
    StatFoldersValueAddress = ADDR_STAT_FOLDERS_VAL
End Property

Public Property Get StatSizeValueAddress() As String
    StatSizeValueAddress = ADDR_STAT_SIZE_VAL
End Property

Public Property Get CriteriaFirstRow() As Long
    CriteriaFirstRow = CRITERIA_FIRST_ROW
End Property

Public Property Get CriteriaLastRow() As Long
    CriteriaLastRow = CRITERIA_LAST_ROW
End Property

Public Property Get CriteriaColTerm1() As Long
    CriteriaColTerm1 = CRITERIA_COL_TERM1
End Property

Public Property Get CriteriaColOp() As Long
    CriteriaColOp = CRITERIA_COL_OP
End Property

Public Property Get CriteriaColTerm2() As Long
    CriteriaColTerm2 = CRITERIA_COL_TERM2
End Property

Public Property Get CriteriaColSizeOp() As Long
    CriteriaColSizeOp = CRITERIA_COL_SIZE_OP
End Property

Public Property Get CriteriaColSizeMb() As Long
    CriteriaColSizeMb = CRITERIA_COL_SIZE_MB
End Property

Public Property Get FileSearchAddress() As String
    FileSearchAddress = ADDR_FILE_SEARCH
End Property

Public Property Get FileSearchTerm() As String
    FileSearchTerm = Trim$(CStr(Dashboard.Range(ADDR_FILE_SEARCH).Value & ""))
End Property

Public Property Get CriteriaTerm1() As String
    CriteriaTerm1 = Trim$(CStr(Dashboard.Range(ADDR_TERM1).Value & ""))
End Property

Public Property Get CriteriaOp() As String
    CriteriaOp = UCase$(Trim$(CStr(Dashboard.Range(ADDR_OP).Value & "")))
End Property

Public Property Get CriteriaTerm2() As String
    CriteriaTerm2 = Trim$(CStr(Dashboard.Range(ADDR_TERM2).Value & ""))
End Property

Public Property Get CriteriaSizeOp() As String
    CriteriaSizeOp = Trim$(CStr(Dashboard.Range(ADDR_SIZE_OP).Value & ""))
End Property

Public Property Get CriteriaSizeMbValue() As Variant
    CriteriaSizeMbValue = Dashboard.Range(ADDR_SIZE_MB).Value
End Property

' Back-compat alias: primary box is File Search (A5).
Public Property Get SearchTermAddress() As String
    SearchTermAddress = ADDR_FILE_SEARCH
End Property

Public Property Get SearchTerm() As String
    SearchTerm = FileSearchTerm
End Property

Public Property Get ResultsHeaderRow() As Long
    ResultsHeaderRow = RESULTS_HEADER_ROW
End Property

Public Property Get ResultsFileCol() As Long
    ResultsFileCol = RESULTS_FILE_COL
End Property

Public Property Get ResultsFileColEnd() As Long
    ResultsFileColEnd = RESULTS_FILE_COL_END
End Property

Public Property Get ResultsTypeCol() As Long
    ResultsTypeCol = RESULTS_TYPE_COL
End Property

Public Property Get ResultsDateCol() As Long
    ResultsDateCol = RESULTS_DATE_COL
End Property

Public Property Get ResultsSizeCol() As Long
    ResultsSizeCol = RESULTS_SIZE_COL
End Property

Public Property Get ResultsDriveCol() As Long
    ResultsDriveCol = RESULTS_DRIVE_COL
End Property

Public Property Get ResultsFolderCol() As Long
    ResultsFolderCol = RESULTS_FOLDER_COL
End Property

Public Property Get ResultsFolderColEnd() As Long
    ResultsFolderColEnd = RESULTS_FOLDER_COL_END
End Property

Public Property Get ResultsLastCol() As Long
    ResultsLastCol = RESULTS_LAST_COL
End Property

Public Property Get ResultsLinkCol() As Long
    ResultsLinkCol = RESULTS_LINK_COL
End Property

Public Property Get ResultsMaxRows() As Long
    ResultsMaxRows = RESULTS_MAX_ROWS
End Property

Public Property Get ResultsFirstDataRow() As Long
    ResultsFirstDataRow = RESULTS_HEADER_ROW + 1
End Property

Public Property Get IngestedHeaderRow() As Long
    IngestedHeaderRow = INGESTED_HEADER_ROW
End Property

Public Property Get IngestedStartCol() As Long
    IngestedStartCol = INGESTED_START_COL
End Property

Public Property Get IngestedColCount() As Long
    IngestedColCount = INGESTED_COL_COUNT
End Property

Public Property Get IngestSummaryScansLabel() As String
    IngestSummaryScansLabel = ADDR_SUM_SCANS_LBL
End Property
Public Property Get IngestSummaryScansValue() As String
    IngestSummaryScansValue = ADDR_SUM_SCANS_VAL
End Property
Public Property Get IngestSummaryFilesLabel() As String
    IngestSummaryFilesLabel = ADDR_SUM_FILES_LBL
End Property
Public Property Get IngestSummaryFilesValue() As String
    IngestSummaryFilesValue = ADDR_SUM_FILES_VAL
End Property
Public Property Get IngestSummaryFoldersLabel() As String
    IngestSummaryFoldersLabel = ADDR_SUM_FOLDERS_LBL
End Property
Public Property Get IngestSummaryFoldersValue() As String
    IngestSummaryFoldersValue = ADDR_SUM_FOLDERS_VAL
End Property
Public Property Get IngestSummarySizeLabel() As String
    IngestSummarySizeLabel = ADDR_SUM_SIZE_LBL
End Property
Public Property Get IngestSummarySizeValue() As String
    IngestSummarySizeValue = ADDR_SUM_SIZE_VAL
End Property

Public Sub EnsureModernSheetNames()
    RenameIfExists "HomePage", SHEET_DASHBOARD
    RenameIfExists "Feed The Beast", SHEET_INGESTION
    RenameIfExists "Data", SHEET_DATABASE
End Sub

' Re-apply known Dashboard chrome after search.
' Never touch criteria / File Search merges — those are user-owned.
Public Sub EnsureDashboardChromeMerges()
    Dim ws As Worksheet
    Dim divider As Range

    On Error Resume Next
    Set ws = Dashboard
    If ws Is Nothing Then Exit Sub

    ' Grey divider (between options and results headers)
    Set divider = ws.Range(ws.Cells(RESULTS_DIVIDER_ROW, 1), ws.Cells(RESULTS_DIVIDER_ROW, RESULTS_LAST_COL))
    divider.UnMerge
    divider.Interior.Color = RGB(166, 166, 166)
    divider.Value2 = Empty

    ' Search stats (labels/values live in merged M:N)
    SafeMergeRange ws, "M3:N3"
    SafeMergeRange ws, "M4:N4"
    SafeMergeRange ws, "M5:N5"

    ' Drive dropdown chrome (File Search A5:G5 is user-owned)
    SafeMergeRange ws, "H5:I5"

    On Error GoTo 0
End Sub

' Lazy build of results band (merges + stripes) when missing — used by search, not Workbook_Open.
Public Sub EnsureResultsDisplayGrid()
    Dim ws As Worksheet
    Dim firstRow As Long
    Dim lastRow As Long

    Set ws = Dashboard
    If ws Is Nothing Then Exit Sub

    UnprotectDashboard
    ClearResultsAutoFilter ws

    firstRow = RESULTS_HEADER_ROW + 1
    lastRow = firstRow + RESULTS_MAX_ROWS - 1
    EnsureLinkColumnHidden ws

    If ResultsGridLooksReady(ws, firstRow, lastRow) Then
        Debug.Print "INDEX CACHE: results grid already ready rows=" & CStr(RESULTS_MAX_ROWS)
        EnsureResultsAutoFilter ws
        Exit Sub
    End If

    BuildResultsDisplayGridCore ws, firstRow, lastRow
End Sub

Private Sub BuildResultsDisplayGridCore(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    Dim r As Long
    Dim i As Long
    Dim rowRng As Range
    Dim stripeGrey As Long
    Dim stripeWhite As Long
    Dim t0 As Double
    Dim screenPrev As Boolean
    Dim eventsPrev As Boolean
    Dim calcPrev As XlCalculation

    t0 = Timer
    screenPrev = Application.ScreenUpdating
    eventsPrev = Application.EnableEvents
    calcPrev = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    On Error Resume Next
    Application.StatusBar = "LAN Search: building results grid (" & CStr(RESULTS_MAX_ROWS) & " rows)..."
    On Error GoTo 0

    EnsureDashboardChromeMerges

    stripeWhite = RGB(255, 255, 255)
    stripeGrey = RGB(242, 242, 242)

    For r = firstRow To lastRow
        i = r - firstRow + 1
        SafeMerge ws, r, RESULTS_FILE_COL, RESULTS_FILE_COL_END
        SafeMerge ws, r, RESULTS_FOLDER_COL, RESULTS_FOLDER_COL_END
        Set rowRng = ws.Range(ws.Cells(r, RESULTS_FILE_COL), ws.Cells(r, RESULTS_LAST_COL))
        If (i Mod 2) = 0 Then
            rowRng.Interior.Color = stripeGrey
        Else
            rowRng.Interior.Color = stripeWhite
        End If
        ' Format per row — never Range(A8:A20007).* (Excel collapses A:K into A8:A20007)
        With ws.Cells(r, RESULTS_FILE_COL)
            .NumberFormat = "@"
            .HorizontalAlignment = xlLeft
            .Font.Color = RGB(5, 99, 193)
            .Font.Underline = xlUnderlineStyleSingle
        End With
        With ws.Range(ws.Cells(r, RESULTS_DRIVE_COL), ws.Cells(r, RESULTS_FOLDER_COL_END))
            .HorizontalAlignment = xlLeft
            .Font.ColorIndex = xlColorIndexAutomatic
            .Font.Underline = xlUnderlineStyleNone
        End With
        If i Mod 250 = 0 Then
            On Error Resume Next
            Application.StatusBar = "LAN Search: results grid " & CStr(i) & "/" & CStr(RESULTS_MAX_ROWS)
            On Error GoTo 0
            DoEvents
        End If
    Next r

    With ws.Range(ws.Cells(firstRow, RESULTS_SIZE_COL), ws.Cells(lastRow, RESULTS_SIZE_COL))
        .NumberFormat = "0.00"
        .HorizontalAlignment = xlRight
    End With
    With ws.Range(ws.Cells(firstRow, RESULTS_DATE_COL), ws.Cells(lastRow, RESULTS_DATE_COL))
        .NumberFormat = "yyyy-mmm-dd"
    End With
    ws.Range(ws.Cells(firstRow, RESULTS_TYPE_COL), ws.Cells(lastRow, RESULTS_TYPE_COL)).HorizontalAlignment = xlCenter

    EnsureLinkColumnHidden ws
    EnsureResultsAutoFilter ws

    On Error Resume Next
    Application.StatusBar = False
    Application.ScreenUpdating = screenPrev
    Application.EnableEvents = eventsPrev
    Application.Calculation = calcPrev
    On Error GoTo 0

    ProtectDashboardResults

    Debug.Print "INDEX CACHE: results grid built rows=" & CStr(RESULTS_MAX_ROWS) & _
                " " & Format$(Timer - t0, "0.000") & "s"
End Sub

' Permanent AutoFilter on Type..Drive headers only (L:O).
' Never attach filter to File A:K / Folder P:T — Excel corrupts those row merges into A8:A20007.
Private Sub EnsureResultsAutoFilter(ByVal ws As Worksheet)
    Dim hdr As Range
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    If ws.AutoFilterMode Then
        ' Already filtered on some range — leave layout alone
        On Error GoTo 0
        Exit Sub
    End If
    Set hdr = ws.Range(ws.Cells(RESULTS_HEADER_ROW, RESULTS_TYPE_COL), _
                       ws.Cells(RESULTS_HEADER_ROW, RESULTS_DRIVE_COL))
    hdr.AutoFilter
    On Error GoTo 0
End Sub

' Undo any active results filter so clears see all rows (does not re-apply filter).
Public Sub ClearResultsAutoFilter(Optional ByVal ws As Worksheet = Nothing)
    If ws Is Nothing Then Set ws = Dashboard
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    If ws.FilterMode Then ws.ShowAllData
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    On Error GoTo 0
End Sub

' On workbook open: drop leftover filter + prior search hits only.
' Never rebuild merges, stripes, fonts, or alignment on the results table.
Public Sub ResetResultsAreaOnOpen()
    Dim ws As Worksheet
    Dim firstRow As Long
    Dim lastRow As Long

    Set ws = Dashboard
    If ws Is Nothing Then Exit Sub

    UnprotectDashboard
    firstRow = RESULTS_HEADER_ROW + 1
    lastRow = firstRow + RESULTS_MAX_ROWS - 1

    ClearResultsAutoFilter ws
    ClearResultsDataArea ws, firstRow, lastRow
    EnsureResultsAutoFilter ws
    Debug.Print "INDEX CACHE: ResetResultsAreaOnOpen cleared filter + results"
End Sub

' Public clear used by search — same merge-safe path as open reset.
Public Sub ClearResultsValues()
    Dim ws As Worksheet
    Dim firstRow As Long
    Dim lastRow As Long
    Set ws = Dashboard
    If ws Is Nothing Then Exit Sub
    firstRow = RESULTS_HEADER_ROW + 1
    lastRow = firstRow + RESULTS_MAX_ROWS - 1
    ClearResultsDataArea ws, firstRow, lastRow
End Sub

' Clear values without touching Interior / Font / MergeArea.
' Bulk ClearContents / Hyperlinks.Delete on A:T spanning row-merges can collapse File into A8:A20007.
Private Sub ClearResultsDataArea(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long)
    On Error Resume Next
    ' Mid columns are single cells — safe to clear as a block
    ws.Range(ws.Cells(firstRow, RESULTS_TYPE_COL), ws.Cells(lastRow, RESULTS_DRIVE_COL)).ClearContents
    ' File / Folder: clear top-left of each row merge only (Value2, no format ops)
    ws.Range(ws.Cells(firstRow, RESULTS_FILE_COL), ws.Cells(lastRow, RESULTS_FILE_COL)).Value2 = Empty
    ws.Range(ws.Cells(firstRow, RESULTS_FOLDER_COL), ws.Cells(lastRow, RESULTS_FOLDER_COL)).Value2 = Empty
    ws.Range(ws.Cells(firstRow, RESULTS_LINK_COL), ws.Cells(lastRow, RESULTS_LINK_COL)).ClearContents
    ws.Range(ADDR_STAT_FILES_VAL).NumberFormat = "#,##0"
    ws.Range(ADDR_STAT_FOLDERS_VAL).NumberFormat = "#,##0"
    ws.Range(ADDR_STAT_SIZE_VAL).NumberFormat = "#,##0.00"
    ws.Range(ADDR_STAT_FILES_VAL).Value = 0
    ws.Range(ADDR_STAT_FOLDERS_VAL).Value = 0
    ws.Range(ADDR_STAT_SIZE_VAL).Value = 0
    On Error GoTo 0
End Sub

Private Sub EnsureLinkColumnHidden(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Columns(RESULTS_LINK_COL).Hidden = True
    On Error GoTo 0
End Sub

' Allow VBA writes / grid build.
Public Sub UnprotectDashboard()
    UnprotectSheet Dashboard
End Sub

Public Sub UnprotectWorkbookForWrite()
    Dim ws As Worksheet
    UnprotectDashboard
    On Error Resume Next
    Set ws = Ingestion
    UnprotectSheet ws
    Set ws = Database
    UnprotectSheet ws
    ' Briefly visible so ListObject / bulk writes are reliable
    ws.Visible = xlSheetVisible
    On Error GoTo 0
End Sub

Private Sub UnprotectSheet(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ws.Unprotect Password:=vbNullString
    ws.Unprotect
    On Error GoTo 0
End Sub

' Full UI lock: File Search A5 + criteria row3 + Drive G5.
' Ingestion locked; Database protected + xlSheetVeryHidden.
' VBA ingestion/search still works via UnprotectWorkbookForWrite / UserInterfaceOnly.
Public Sub ProtectDashboardResults()
    ProtectWorkbookUi
End Sub

Public Sub ProtectWorkbookUi()
    Dim was As Worksheet
    On Error Resume Next
    Set was = mWorkbook.ActiveSheet
    On Error GoTo 0

    ProtectDashboardSheet
    ProtectIngestionSheet
    ProtectAndHideDatabase

    On Error Resume Next
    If Not was Is Nothing Then
        If was.Visible = xlSheetVisible Then was.Activate
    End If
    On Error GoTo 0
    Debug.Print "INDEX CACHE: workbook UI locked (A5:G5 File Search + F3/G3 size + H5 drive)"
End Sub

' After search: stay on Dashboard and show results start. Does NOT touch FreezePanes.
Public Sub FocusDashboardResultsHome()
    Dim ws As Worksheet
    Dim screenPrev As Boolean

    On Error Resume Next
    Set ws = Dashboard
    If ws Is Nothing Then Exit Sub
    screenPrev = Application.ScreenUpdating
    Application.ScreenUpdating = False
    ws.Activate
    Application.Goto Reference:=ws.Range("A8"), Scroll:=True
    Application.ScreenUpdating = screenPrev
    On Error GoTo 0
End Sub

Private Sub ProtectDashboardSheet()
    Dim ws As Worksheet

    Set ws = Dashboard
    If ws Is Nothing Then Exit Sub

    UnprotectSheet ws
    On Error Resume Next

    ' Cells.Locked=True on every protect would re-lock merged results and bring
    ' back Excel's protect dialog on click. Do the lock map once per session.
    If Not gResultsCopyUnlockDone Then
        ws.Cells.Locked = True
        ' Free-text: Contains A3:B3, Contains D3:E3, Size MB G3, File Search A5:F5
        UnlockMergeOrRange ws, ADDR_TERM1
        UnlockMergeOrRange ws, ADDR_TERM2
        UnlockMergeOrRange ws, ADDR_SIZE_MB
        UnlockMergeOrRange ws, ADDR_FILE_SEARCH
        ' Dropdowns: Op C3, Size is F3, Flexible H3, Folders I3, Drive H5
        UnlockMergeOrRange ws, ADDR_OP
        UnlockMergeOrRange ws, ADDR_SIZE_OP
        UnlockMergeOrRange ws, ADDR_FLEXIBLE_MATCH
        UnlockMergeOrRange ws, ADDR_SEARCH_FOLDERS
        UnlockMergeOrRange ws, ADDR_DRIVE_FILTER
        ws.Columns(RESULTS_LINK_COL).Locked = False
        EnsureLinkColumnHidden ws
        UnlockResultsDataArea ws
        gResultsCopyUnlockDone = True
    Else
        UnlockMergeOrRange ws, ADDR_TERM1
        UnlockMergeOrRange ws, ADDR_TERM2
        UnlockMergeOrRange ws, ADDR_SIZE_MB
        UnlockMergeOrRange ws, ADDR_FILE_SEARCH
        UnlockMergeOrRange ws, ADDR_OP
        UnlockMergeOrRange ws, ADDR_SIZE_OP
        UnlockMergeOrRange ws, ADDR_FLEXIBLE_MATCH
        UnlockMergeOrRange ws, ADDR_SEARCH_FOLDERS
        UnlockMergeOrRange ws, ADDR_DRIVE_FILTER
        EnsureLinkColumnHidden ws
    End If

    ApplyCriteriaOperatorValidation
    ApplyOptionDropdownValidation

    ws.Columns(RESULTS_SIZE_COL).ColumnWidth = 10
    CaptureDashboardFixedColumnWidthsIfNeeded ws
    modEntryPoints.SetDashboardFixedWidth RESULTS_SIZE_COL, 10
    RestoreDashboardFixedColumnWidths ws

    ' Filter only if missing — do not tear down/rebuild (corrupts File row merges)
    EnsureResultsAutoFilter ws

    ApplySheetProtect ws, True, True
    ws.EnableSelection = xlNoRestrictions
    On Error GoTo 0
End Sub

Private Sub UnlockResultsDataArea(ByVal ws As Worksheet)
    Dim r As Long
    Dim firstRow As Long
    Dim lastRow As Long

    If ws Is Nothing Then Exit Sub
    firstRow = RESULTS_HEADER_ROW + 1
    lastRow = firstRow + RESULTS_MAX_ROWS - 1
    On Error Resume Next
    ' Non-merged Type..Drive can be a block
    ws.Range(ws.Cells(firstRow, RESULTS_TYPE_COL), ws.Cells(lastRow, RESULTS_DRIVE_COL)).Locked = False
    ws.Columns(RESULTS_LINK_COL).Locked = False
    ' Per-row only — a tall A8:A20007 / P8:P20007 unlock collapses the merges
    For r = firstRow To lastRow
        ws.Cells(r, RESULTS_FILE_COL).MergeArea.Locked = False
        ws.Cells(r, RESULTS_FOLDER_COL).MergeArea.Locked = False
        If r Mod 2000 = 0 Then
            Application.StatusBar = "LAN Search: enabling results copy " & CStr(r - firstRow + 1) & "/" & CStr(RESULTS_MAX_ROWS)
            DoEvents
        End If
    Next r
    Application.StatusBar = False
    On Error GoTo 0
End Sub

Private Sub UnlockMergeOrRange(ByVal ws As Worksheet, ByVal addr As String)
    Dim rng As Range
    On Error Resume Next
    Set rng = ws.Range(addr)
    If rng Is Nothing Then Exit Sub
    If rng.MergeCells Then
        rng.MergeArea.Locked = False
    Else
        rng.Locked = False
    End If
    On Error GoTo 0
End Sub

Private Sub ProtectIngestionSheet()
    Dim ws As Worksheet
    Set ws = Ingestion
    If ws Is Nothing Then Exit Sub
    UnprotectSheet ws
    On Error Resume Next
    ws.Cells.Locked = True
    ApplySheetProtect ws, False
    ws.EnableSelection = xlNoRestrictions
    On Error GoTo 0
End Sub

Private Sub ProtectAndHideDatabase()
    Dim ws As Worksheet
    Set ws = Database
    If ws Is Nothing Then Exit Sub
    UnprotectSheet ws
    On Error Resume Next
    ws.Cells.Locked = True
    ApplySheetProtect ws, False
    ws.EnableSelection = xlNoRestrictions
    ws.Visible = xlSheetVeryHidden ' 2 — not listed in Unhide UI
    On Error GoTo 0
End Sub

Private Sub ApplySheetProtect(ByVal ws As Worksheet, Optional ByVal allowFormattingColumns As Boolean = False, _
                              Optional ByVal allowFiltering As Boolean = False)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ws.Protect Password:=vbNullString, _
               DrawingObjects:=False, _
               Contents:=True, _
               Scenarios:=True, _
               UserInterfaceOnly:=True, _
               AllowFormattingCells:=False, _
               AllowFormattingColumns:=allowFormattingColumns, _
               AllowFormattingRows:=False, _
               AllowInsertingColumns:=False, _
               AllowInsertingRows:=False, _
               AllowInsertingHyperlinks:=False, _
               AllowDeletingColumns:=False, _
               AllowDeletingRows:=False, _
               AllowSorting:=allowFiltering, _
               AllowFiltering:=allowFiltering, _
               AllowUsingPivotTables:=False
    On Error GoTo 0
End Sub

' Fixed chrome columns (not user-resizable): A:G, H:J, L:N
' G = Size MB (criteria) — snap back. K has no chrome — leave resizable.
Public Sub CaptureDashboardFixedColumnWidthsIfNeeded(ByVal ws As Worksheet)
    Dim c As Variant
    If ws Is Nothing Then Exit Sub
    If modEntryPoints.DashboardFixedWidthsReady Then Exit Sub
    On Error Resume Next
    For Each c In Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 13, 14)
        modEntryPoints.SetDashboardFixedWidth CLng(c), CDbl(ws.Columns(CLng(c)).ColumnWidth)
    Next c
    modEntryPoints.DashboardFixedWidthsReady = True
    On Error GoTo 0
End Sub

Public Sub RestoreDashboardFixedColumnWidths(ByVal ws As Worksheet)
    Dim c As Variant
    Dim w As Double
    If ws Is Nothing Then Exit Sub
    If Not modEntryPoints.DashboardFixedWidthsReady Then Exit Sub
    On Error Resume Next
    For Each c In Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 12, 13, 14)
        w = modEntryPoints.GetDashboardFixedWidth(CLng(c))
        If w > 0 Then
            If Abs(ws.Columns(CLng(c)).ColumnWidth - w) > 0.05 Then
                ws.Columns(CLng(c)).ColumnWidth = w
            End If
        End If
    Next c
    On Error GoTo 0
End Sub

Private Function ResultsGridLooksReady(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long) As Boolean
    Dim needFileCols As Long
    Dim needFolderCols As Long
    Dim ma As Range

    On Error Resume Next
    ResultsGridLooksReady = False
    needFileCols = RESULTS_FILE_COL_END - RESULTS_FILE_COL + 1
    needFolderCols = RESULTS_FOLDER_COL_END - RESULTS_FOLDER_COL + 1

    ' Require single-row A:K / P:T merges — a tall A8:K20000 still has Columns=11 and used to look "ready".
    Set ma = ws.Cells(firstRow, RESULTS_FILE_COL).MergeArea
    If ma.Columns.Count < needFileCols Or ma.Rows.Count <> 1 Then Exit Function
    Set ma = ws.Cells(lastRow, RESULTS_FILE_COL).MergeArea
    If ma.Columns.Count < needFileCols Or ma.Rows.Count <> 1 Then Exit Function
    Set ma = ws.Cells(firstRow, RESULTS_FOLDER_COL).MergeArea
    If ma.Columns.Count < needFolderCols Or ma.Rows.Count <> 1 Then Exit Function
    Set ma = ws.Cells(lastRow, RESULTS_FOLDER_COL).MergeArea
    If ma.Columns.Count < needFolderCols Or ma.Rows.Count <> 1 Then Exit Function

    ResultsGridLooksReady = True
End Function

Private Sub SafeMerge(ByVal ws As Worksheet, ByVal row As Long, ByVal c1 As Long, ByVal c2 As Long)
    Dim rng As Range
    If ws Is Nothing Then Exit Sub
    If c2 < c1 Then Exit Sub
    On Error Resume Next
    Set rng = ws.Range(ws.Cells(row, c1), ws.Cells(row, c2))
    If rng.MergeCells Then
        If rng.MergeArea.Address = rng.Address Then Exit Sub
    End If
    rng.UnMerge
    rng.Merge
    On Error GoTo 0
End Sub

Private Sub SafeMergeRange(ByVal ws As Worksheet, ByVal addr As String)
    Dim rng As Range
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    Set rng = ws.Range(addr)
    If Not rng.MergeCells Then
        rng.Merge
    End If
    On Error GoTo 0
End Sub

Public Sub ClearLegacyCriteriaValidation()
    Dim ws As Worksheet

    Set ws = Dashboard
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ' Criteria + File Search only (A3:G5). Never touch Drive H5:I5 — ProtectWorkbookUi
    ' used to wipe the AccDB/sheet drive dropdown every lock.
    ws.Range("A3:G5").Validation.Delete
    On Error GoTo 0
End Sub

Public Sub ApplyCriteriaOperatorValidation()
    Dim rng As Range

    ClearLegacyCriteriaValidation

    On Error Resume Next
    Set rng = Dashboard.Range(ADDR_OP)
    If rng.MergeCells Then Set rng = rng.MergeArea.Cells(1, 1)
    rng.Validation.Delete
    On Error GoTo 0
    With rng.Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
             Formula1:="AND,OR,NOT"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = False
        .ShowError = True
        .ErrorTitle = "Operator"
        .ErrorMessage = "Pick AND, OR, or NOT from the list (or leave blank)."
    End With

    On Error Resume Next
    Set rng = Dashboard.Range(ADDR_SIZE_OP)
    If rng.MergeCells Then Set rng = rng.MergeArea.Cells(1, 1)
    rng.Validation.Delete
    On Error GoTo 0
    With rng.Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
             Formula1:="Over,Under"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = False
        .ShowError = True
        .ErrorTitle = "Size"
        .ErrorMessage = "Pick Over or Under from the list (or leave blank)."
    End With

    ' Migrate leftover < / > from older workbooks
    On Error Resume Next
    Select Case UCase$(Trim$(CStr(rng.Value & "")))
        Case ">", "GT", "GREATER THAN": rng.Value = "Over"
        Case "<", "LT", "LESS THAN": rng.Value = "Under"
    End Select
    On Error GoTo 0
End Sub

' Flexible match / Search folders — list pick only (Stop blocks free typing).
Public Sub ApplyOptionDropdownValidation()
    Dim rng As Range

    On Error Resume Next
    Set rng = Dashboard.Range(ADDR_FLEXIBLE_MATCH)
    If rng.MergeCells Then Set rng = rng.MergeArea.Cells(1, 1)
    rng.Validation.Delete
    On Error GoTo 0
    With rng.Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
             Formula1:="Yes,No"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = False
        .ShowError = True
        .ErrorTitle = "Flexible match"
        .ErrorMessage = "Pick Yes or No from the list."
    End With

    On Error Resume Next
    Set rng = Dashboard.Range(ADDR_SEARCH_FOLDERS)
    If rng.MergeCells Then Set rng = rng.MergeArea.Cells(1, 1)
    rng.Validation.Delete
    On Error GoTo 0
    With rng.Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
             Formula1:="Yes,No,Only"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = False
        .ShowError = True
        .ErrorTitle = "Search folders"
        .ErrorMessage = "Pick Yes, No, or Only from the list."
    End With
End Sub

Private Sub RenameIfExists(ByVal oldName As String, ByVal newName As String)
    Dim ws As Worksheet
    Dim existing As Worksheet

    On Error Resume Next
    Set ws = mWorkbook.Worksheets(oldName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    Set existing = mWorkbook.Worksheets(newName)
    On Error GoTo 0
    If Not existing Is Nothing Then Exit Sub

    ws.Name = newName
End Sub
