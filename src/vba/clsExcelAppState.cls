VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsExcelAppState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mScreenUpdating As Boolean
Private mEnableEvents As Boolean
Private mDisplayAlerts As Boolean
Private mStatusBar As Variant
Private mStatusBarIsDefault As Boolean
Private mCaptured As Boolean

Public Sub Capture()
    mScreenUpdating = Application.ScreenUpdating
    mEnableEvents = Application.EnableEvents
    mDisplayAlerts = Application.DisplayAlerts
    ' Application.StatusBar is False when Excel owns the bar — never restore that as text "FALSE"
    If VarType(Application.StatusBar) = vbBoolean Then
        mStatusBarIsDefault = True
        mStatusBar = False
    Else
        mStatusBarIsDefault = (Len(CStr(Application.StatusBar & "")) = 0)
        mStatusBar = Application.StatusBar
    End If
    mCaptured = True
End Sub

Public Sub BeginFast()
    If Not mCaptured Then Capture
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
End Sub

Public Sub SetStatus(ByVal message As String)
    On Error Resume Next
    Application.StatusBar = message
    On Error GoTo 0
End Sub

Public Sub ClearStatus()
    ' Avoid leaving the literal word "FALSE" in the status bar (Excel quirk).
    On Error Resume Next
    Application.DisplayStatusBar = True
    Application.StatusBar = vbNullString
    Application.StatusBar = False
    On Error GoTo 0
End Sub

Public Sub Restore()
    On Error Resume Next
    If mCaptured Then
        Application.ScreenUpdating = mScreenUpdating
        Application.EnableEvents = mEnableEvents
        Application.DisplayAlerts = mDisplayAlerts
        If mStatusBarIsDefault Or VarType(mStatusBar) = vbBoolean Then
            Application.StatusBar = vbNullString
            Application.StatusBar = False
        ElseIf Len(CStr(mStatusBar & "")) = 0 Then
            Application.StatusBar = vbNullString
            Application.StatusBar = False
        Else
            Application.StatusBar = CStr(mStatusBar)
        End If
    Else
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.DisplayAlerts = True
        Application.StatusBar = vbNullString
        Application.StatusBar = False
    End If
    On Error GoTo 0
End Sub
