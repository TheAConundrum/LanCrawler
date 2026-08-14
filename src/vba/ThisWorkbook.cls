VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' Document module — injected into the workbook's ThisWorkbook.
' VBA cannot true-background load; OnTime defers cache warm so Open returns
' immediately and the user can work while the index loads a moment later.

Private mWarmDue As Date

Private Sub Workbook_Open()
    On Error Resume Next
    Application.EnableEvents = True
    Debug.Print "LAN Search: Workbook_Open — scheduling cache warm"
    mWarmDue = Now + TimeValue("00:00:02")
    Application.OnTime EarliestTime:=mWarmDue, Procedure:="WarmIndexCache", Schedule:=True
    If Err.Number <> 0 Then
        Debug.Print "LAN Search: OnTime schedule failed Err=" & CStr(Err.Number) & " " & Err.Description
    End If
    On Error GoTo 0
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    If mWarmDue <> 0 Then
        Application.OnTime EarliestTime:=mWarmDue, Procedure:="WarmIndexCache", Schedule:=False
    End If
    On Error GoTo 0
End Sub

' Open containing folder on demand — avoids Hyperlinks.Add at search time.
Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    On Error Resume Next
    HandleDashboardResultDoubleClick Sh, Target, Cancel
    On Error GoTo 0
End Sub

' Snap chrome columns back if the user dragged them (Excel cannot lock widths per-column).
Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    Dim map As clsSheetMap
    On Error Resume Next
    If Sh Is Nothing Then Exit Sub
    If StrComp(Sh.Name, "Dashboard", vbTextCompare) <> 0 Then Exit Sub
    If Not DashboardFixedWidthsReady Then Exit Sub
    Set map = New clsSheetMap
    map.Init Sh.Parent
    map.RestoreDashboardFixedColumnWidths Sh
    On Error GoTo 0
End Sub

' Unlocked results are selectable/copyable; revert typed/pasted edits.
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    RejectDashboardResultsEdit Sh, Target
    On Error GoTo 0
End Sub
