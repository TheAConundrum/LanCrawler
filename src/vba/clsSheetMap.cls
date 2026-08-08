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
' Row 1 = instructions, row 2 = criteria headers, rows 3-5 = criteria,
' Options: H3 / I3:J3 labels, H4 / I4:J4 values; Drive dropdown H5;
' Search stats values M3 / M4 / M5 (merged M:N); row 6 divider;
' Results headers row 7; data row 8+.
Private Const CRITERIA_HEADER_ROW As Long = 2
Private Const CRITERIA_FIRST_ROW As Long = 3
Private Const CRITERIA_LAST_ROW As Long = 5
Private Const CRITERIA_COL_TERM1 As Long = 2 ' B
Private Const CRITERIA_COL_OP As Long = 3    ' C
Private Const CRITERIA_COL_TERM2 As Long = 4 ' D
Private Const CRITERIA_COL_SIZE_OP As Long = 5 ' E  < or >
Private Const CRITERIA_COL_SIZE_MB As Long = 6 ' F
Private Const ADDR_FLEXIBLE_MATCH As String = "H4" ' Yes/No (label at H3)
Private Const ADDR_SEARCH_FOLDERS As String = "I4" ' Yes/No/Only (label at I3:J3)
Private Const ADDR_DRIVE_FILTER As String = "H5"   ' Drive dropdown value (merged H5:J5 ok)
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
Private Const RESULTS_MAX_ROWS As Long = 5000

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

' Backward-compatible aliases used by older call sites
Public Property Get ResultsCol() As Long
    ResultsCol = RESULTS_FILE_COL
End Property

Public Property Get ResultsColCount() As Long
    ResultsColCount = RESULTS_LAST_COL - RESULTS_FILE_COL + 1
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

' Intentionally empty: Dashboard / Ingestion chrome is user-owned.
' Not called from RunSearch / RunIndex (kept only so old call sites are harmless).
Public Sub EnsureLayoutLabels()
End Sub

' Re-apply known Dashboard chrome after search.
' Never touch criteria (2–5) or the results header row — those are user-owned.
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

    ' Drive / folder / flexible match chrome
    SafeMergeRange ws, "H5:J5"
    SafeMergeRange ws, "I3:J3"
    SafeMergeRange ws, "I4:J4"

    On Error GoTo 0
End Sub

' One-time (or repair) build of the fixed results band: merges + grey/white stripes.
' Searches then only ClearContents / Value2 / hyperlinks into top-left merge cells.
Public Sub EnsureResultsDisplayGrid()
    Dim ws As Worksheet
    Dim firstRow As Long
    Dim lastRow As Long
    Dim r As Long
    Dim i As Long
    Dim rowRng As Range
    Dim stripeGrey As Long
    Dim stripeWhite As Long
    Dim t0 As Double
    Dim screenPrev As Boolean
    Dim eventsPrev As Boolean

    Set ws = Dashboard
    If ws Is Nothing Then Exit Sub

    UnprotectDashboard

    firstRow = RESULTS_HEADER_ROW + 1
    lastRow = firstRow + RESULTS_MAX_ROWS - 1
    EnsureLinkColumnHidden ws

    If ResultsGridLooksReady(ws, firstRow, lastRow) Then
        Debug.Print "INDEX CACHE: results grid already ready rows=" & CStr(RESULTS_MAX_ROWS)
        ProtectDashboardResults
        Exit Sub
    End If

    t0 = Timer
    screenPrev = Application.ScreenUpdating
    eventsPrev = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False
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
        If i Mod 250 = 0 Then DoEvents
    Next r

    With ws.Range(ws.Cells(firstRow, RESULTS_SIZE_COL), ws.Cells(lastRow, RESULTS_SIZE_COL))
        .NumberFormat = "0.00"
        .HorizontalAlignment = xlRight
    End With
    With ws.Range(ws.Cells(firstRow, RESULTS_DATE_COL), ws.Cells(lastRow, RESULTS_DATE_COL))
        .NumberFormat = "yyyy-mmm-dd"
    End With
    ws.Range(ws.Cells(firstRow, RESULTS_TYPE_COL), ws.Cells(lastRow, RESULTS_TYPE_COL)).HorizontalAlignment = xlCenter

    ' File column looks clickable (actual open is double-click → hidden UNC col)
    With ws.Range(ws.Cells(firstRow, RESULTS_FILE_COL), ws.Cells(lastRow, RESULTS_FILE_COL))
        .Font.Color = RGB(5, 99, 193)
        .Font.Underline = xlUnderlineStyleSingle
    End With

    EnsureLinkColumnHidden ws

    On Error Resume Next
    Application.StatusBar = False
    Application.ScreenUpdating = screenPrev
    Application.EnableEvents = eventsPrev
    On Error GoTo 0

    ProtectDashboardResults

    Debug.Print "INDEX CACHE: results grid built rows=" & CStr(RESULTS_MAX_ROWS) & _
                " " & Format$(Timer - t0, "0.000") & "s"
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

' Full UI lock: only Key Word + Size MB cells editable on Dashboard.
' Ingestion locked; Database protected + xlSheetVeryHidden.
' VBA ingestion/search still works via UnprotectWorkbookForWrite / UserInterfaceOnly.
Public Sub ProtectDashboardResults()
    ProtectWorkbookUi
End Sub

Public Sub ProtectWorkbookUi()
    ProtectDashboardSheet
    ProtectIngestionSheet
    ProtectAndHideDatabase
    Debug.Print "INDEX CACHE: workbook UI locked (keywords+MB only; Database very hidden)"
End Sub

Private Sub ProtectDashboardSheet()
    Dim ws As Worksheet

    Set ws = Dashboard
    If ws Is Nothing Then Exit Sub

    UnprotectSheet ws
    On Error Resume Next

    ws.Cells.Locked = True
    ' Free-text search only: Key Word 1, Key Word 2, Size MB (not ops / options / drive)
    ws.Range(ws.Cells(CRITERIA_FIRST_ROW, CRITERIA_COL_TERM1), _
             ws.Cells(CRITERIA_LAST_ROW, CRITERIA_COL_TERM1)).Locked = False
    ws.Range(ws.Cells(CRITERIA_FIRST_ROW, CRITERIA_COL_TERM2), _
             ws.Cells(CRITERIA_LAST_ROW, CRITERIA_COL_TERM2)).Locked = False
    ws.Range(ws.Cells(CRITERIA_FIRST_ROW, CRITERIA_COL_SIZE_MB), _
             ws.Cells(CRITERIA_LAST_ROW, CRITERIA_COL_SIZE_MB)).Locked = False
    ' Hidden UNC targets must be readable while sheet is protected
    ws.Columns(RESULTS_LINK_COL).Locked = False
    EnsureLinkColumnHidden ws

    ApplySheetProtect ws
    ws.EnableSelection = xlNoRestrictions
    On Error GoTo 0
End Sub

Private Sub ProtectIngestionSheet()
    Dim ws As Worksheet
    Set ws = Ingestion
    If ws Is Nothing Then Exit Sub
    UnprotectSheet ws
    On Error Resume Next
    ws.Cells.Locked = True
    ApplySheetProtect ws
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
    ApplySheetProtect ws
    ws.EnableSelection = xlNoRestrictions
    ws.Visible = xlSheetVeryHidden ' 2 — not listed in Unhide UI
    On Error GoTo 0
End Sub

Private Sub ApplySheetProtect(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ws.Protect Password:=vbNullString, _
               DrawingObjects:=False, _
               Contents:=True, _
               Scenarios:=True, _
               UserInterfaceOnly:=True, _
               AllowFormattingCells:=False, _
               AllowFormattingColumns:=False, _
               AllowFormattingRows:=False, _
               AllowInsertingColumns:=False, _
               AllowInsertingRows:=False, _
               AllowInsertingHyperlinks:=False, _
               AllowDeletingColumns:=False, _
               AllowDeletingRows:=False, _
               AllowSorting:=False, _
               AllowFiltering:=False, _
               AllowUsingPivotTables:=False
    On Error GoTo 0
End Sub

Private Function ResultsGridLooksReady(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long) As Boolean
    On Error Resume Next
    ResultsGridLooksReady = False
    If ws.Cells(firstRow, RESULTS_FILE_COL).MergeArea.Columns.Count < (RESULTS_FILE_COL_END - RESULTS_FILE_COL + 1) Then Exit Function
    If ws.Cells(lastRow, RESULTS_FILE_COL).MergeArea.Columns.Count < (RESULTS_FILE_COL_END - RESULTS_FILE_COL + 1) Then Exit Function
    If ws.Cells(firstRow, RESULTS_FOLDER_COL).MergeArea.Columns.Count < (RESULTS_FOLDER_COL_END - RESULTS_FOLDER_COL + 1) Then Exit Function
    If ws.Cells(lastRow, RESULTS_FOLDER_COL).MergeArea.Columns.Count < (RESULTS_FOLDER_COL_END - RESULTS_FOLDER_COL + 1) Then Exit Function
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

Public Sub ApplyCriteriaOperatorValidation()
    Dim rngOp As Range
    Dim rngSize As Range

    Set rngOp = Dashboard.Range(Dashboard.Cells(CRITERIA_FIRST_ROW, CRITERIA_COL_OP), _
                               Dashboard.Cells(CRITERIA_LAST_ROW, CRITERIA_COL_OP))
    On Error Resume Next
    rngOp.Validation.Delete
    On Error GoTo 0
    With rngOp.Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
             Formula1:="AND,OR,NOT"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Operator"
        .InputMessage = "Blank = first keyword only. AND/OR/NOT joins Key Word columns on this row."
        .ShowInput = True
        .ShowError = True
    End With

    Set rngSize = Dashboard.Range(Dashboard.Cells(CRITERIA_FIRST_ROW, CRITERIA_COL_SIZE_OP), _
                                 Dashboard.Cells(CRITERIA_LAST_ROW, CRITERIA_COL_SIZE_OP))
    On Error Resume Next
    rngSize.Validation.Delete
    On Error GoTo 0
    With rngSize.Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
             Formula1:="<,>"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Size"
        .InputMessage = "Optional. Use with Size (MB): e.g. > and 2 means larger than 2 MB."
        .ShowInput = True
        .ShowError = True
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
