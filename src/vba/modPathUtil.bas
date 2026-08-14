Attribute VB_Name = "modPathUtil"
Option Explicit
Option Private Module

' Path / string helpers for LAN Search Tool.
' Option Private Module keeps these out of Excel's Assign Macro list
' while remaining usable by other modules in this workbook.

#If VBA7 Then
Private Declare PtrSafe Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" ( _
    ByVal lpszLocalName As String, _
    ByVal lpszRemoteName As String, _
    ByRef cbRemoteName As Long) As Long
#Else
Private Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" ( _
    ByVal lpszLocalName As String, _
    ByVal lpszRemoteName As String, _
    ByRef cbRemoteName As Long) As Long
#End If

Private Const SEPARATORS_TO_STRIP As String = " .,;-_'()[]{}`~!@#$%^&+=|"
Private Const BYTES_PER_MB As Double = 1048576#
Private Const MIN_SIZE_MB As Double = 0.01

' Docs + images + parent archives (leading/trailing | for safe InStr token match).
' Split volumes (.r01, .z01, .7z.001, .part2.rar, ...) are rejected separately.
Private Const ALLOWED_INDEX_EXTS As String = _
    "|PDF|DOC|DOCX|XLS|XLSX|XLSM|XLSB|PPT|PPTX|TXT|RTF|CSV|ODT|ODS|ODP|MSG|EML|" & _
    "|JPG|JPEG|PNG|GIF|TIF|TIFF|BMP|WEBP|" & _
    "|ZIP|7Z|RAR|TAR|GZ|TGZ|BZ2|XZ|CAB|" & _
    "|MP4|MOV|AVI|MKV|WMV|WEBM|M4V|MPG|MPEG|FLV|3GP|TS|MTS|M2TS|"

Public Function ParentFolderPath(ByVal fullPath As String) As String
    Dim p As String
    Dim slash As Long
    p = Replace(Trim$(fullPath), "/", "\")
    If Len(p) = 0 Then
        ParentFolderPath = vbNullString
        Exit Function
    End If
    If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    slash = InStrRev(p, "\")
    If slash <= 0 Then
        ParentFolderPath = vbNullString
        Exit Function
    End If
    ParentFolderPath = Left$(p, slash) ' includes trailing \
End Function

' Convert raw file bytes to Size MB (2 decimals). Values under 0.01 MB become 0.01.
Public Function BytesToSizeMb(ByVal sizeBytes As Double) As Double
    Dim mb As Double
    If sizeBytes <= 0 Then
        BytesToSizeMb = MIN_SIZE_MB
        Exit Function
    End If
    mb = sizeBytes / BYTES_PER_MB
    If mb < MIN_SIZE_MB Then
        BytesToSizeMb = MIN_SIZE_MB
    Else
        BytesToSizeMb = Round(mb, 2)
    End If
End Function

' Normalize a stored size cell to SizeMB.
' Never use (stored & 0) — that concatenates ("445" & 0 = "4450") and corrupts whole MB values,
' which then looked like "bytes" and displayed as 0.01 while still matching SizeMB>50 in AccDB.
' Legacy SizeBytes migration: only treat values >= 1,000,000 as raw bytes (~1 MB+ in bytes).
Public Function CoerceStoredSizeMb(ByVal stored As Variant) As Double
    Dim v As Double
    On Error Resume Next
    If IsNull(stored) Or IsEmpty(stored) Then
        CoerceStoredSizeMb = MIN_SIZE_MB
        Exit Function
    End If
    If VarType(stored) = vbString Then
        If Len(Trim$(CStr(stored))) = 0 Then
            CoerceStoredSizeMb = MIN_SIZE_MB
            Exit Function
        End If
        v = Val(Replace(CStr(stored), ",", "."))
    Else
        v = CDbl(stored)
    End If
    If Err.Number <> 0 Then
        Err.Clear
        CoerceStoredSizeMb = MIN_SIZE_MB
        Exit Function
    End If
    On Error GoTo 0

    ' >= 1e6 cannot be a realistic single-file SizeMB; treat as legacy SizeBytes
    If v >= 1000000# Then
        CoerceStoredSizeMb = BytesToSizeMb(v)
    ElseIf v <= 0# Then
        CoerceStoredSizeMb = MIN_SIZE_MB
    ElseIf v < MIN_SIZE_MB Then
        CoerceStoredSizeMb = MIN_SIZE_MB
    Else
        CoerceStoredSizeMb = Round(v, 2)
    End If
End Function

Public Function FileExtension(ByVal fileName As String) As String
    FileExtension = ExtensionOnly(fileName)
End Function

Public Function IsAllowedIndexExtension(ByVal fileName As String) As Boolean
    Dim ext As String
    ext = ExtensionOnly(fileName)
    If Len(ext) = 0 Then
        IsAllowedIndexExtension = False
        Exit Function
    End If
    If InStr(1, ALLOWED_INDEX_EXTS, "|" & ext & "|", vbBinaryCompare) = 0 Then
        IsAllowedIndexExtension = False
        Exit Function
    End If
    ' Keep parent archive only (tmllist.rar); skip multi-volume siblings
    If IsSplitArchiveVolume(fileName, ext) Then
        IsAllowedIndexExtension = False
        Exit Function
    End If
    IsAllowedIndexExtension = True
End Function

Public Function IsJunkFolderName(ByVal folderName As String) As Boolean
    Dim n As String
    n = UCase$(Trim$(folderName))
    IsJunkFolderName = ( _
        n = "$RECYCLE.BIN" Or _
        n = "SYSTEM VOLUME INFORMATION" Or _
        n = "THUMBS.DB" Or _
        n = ".GIT" Or _
        n = ".SVN" Or _
        n = ".HG" Or _
        n = "__PYCACHE__" Or _
        n = "NODE_MODULES" Or _
        n = ".VS" Or _
        n = ".IDEA" Or _
        n = "RECYCLER" _
    )
End Function

Public Function IsJunkFileName(ByVal fileName As String) As Boolean
    Dim n As String
    Dim ext As String

    n = Trim$(fileName)
    If Len(n) = 0 Then
        IsJunkFileName = True
        Exit Function
    End If

    ' Office lock / temp siblings
    If Left$(n, 2) = "~$" Then
        IsJunkFileName = True
        Exit Function
    End If

    ext = ExtensionOnly(n)
    n = UCase$(n)

    IsJunkFileName = ( _
        n = "THUMBS.DB" Or _
        n = "DESKTOP.INI" Or _
        n = ".DS_STORE" Or _
        ext = "TMP" Or _
        ext = "TEMP" Or _
        ext = "BAK" Or _
        ext = "LOG" Or _
        ext = "LCK" Or _
        ext = "PART" _
    )
End Function

Private Function ExtensionOnly(ByVal fileName As String) As String
    Dim base As String
    Dim dot As Long

    base = BaseNameUpper(fileName)
    If Len(base) = 0 Then
        ExtensionOnly = vbNullString
        Exit Function
    End If
    dot = InStrRev(base, ".")
    If dot <= 1 Or dot = Len(base) Then
        ExtensionOnly = vbNullString
        Exit Function
    End If
    ExtensionOnly = Mid$(base, dot + 1)
End Function

' True for multi-volume pieces when the parent (.rar / .zip / .7z) is enough.
Private Function IsSplitArchiveVolume(ByVal fileName As String, ByVal ext As String) As Boolean
    Dim base As String
    Dim stem As String
    Dim partTok As String
    Dim partNum As String
    Dim p As Long
    Dim dot As Long

    base = BaseNameUpper(fileName)
    ext = UCase$(Trim$(ext))

    ' Classic RAR volumes: tmllist.r00 .. tmllist.r100 (ext = R## / R###)
    If Left$(ext, 1) = "R" And Len(ext) >= 3 And Len(ext) <= 4 Then
        If IsDigitsOnly(Mid$(ext, 2)) Then
            IsSplitArchiveVolume = True
            Exit Function
        End If
    End If

    ' Split ZIP volumes: name.z01 .. name.z99
    If Left$(ext, 1) = "Z" And Len(ext) = 3 Then
        If IsDigitsOnly(Mid$(ext, 2)) Then
            IsSplitArchiveVolume = True
            Exit Function
        End If
    End If

    ' name.7z.001 / name.zip.001 / name.rar.001
    If IsDigitsOnly(ext) And (Len(ext) = 2 Or Len(ext) = 3) Then
        stem = Left$(base, Len(base) - Len(ext) - 1)
        If Right$(stem, 3) = ".7Z" Or Right$(stem, 4) = ".ZIP" Or Right$(stem, 4) = ".RAR" Then
            IsSplitArchiveVolume = True
            Exit Function
        End If
    End If

    ' New-style: name.part2.rar / name.part02.zip — keep part1 / part01 only
    If ext = "RAR" Or ext = "ZIP" Or ext = "7Z" Then
        p = InStrRev(base, ".PART")
        If p > 0 Then
            partTok = Mid$(base, p + 5) ' after ".PART"
            dot = InStrRev(partTok, ".")
            If dot > 1 Then
                partNum = Left$(partTok, dot - 1)
                If IsDigitsOnly(partNum) Then
                    If CLng(partNum) <> 1 Then
                        IsSplitArchiveVolume = True
                        Exit Function
                    End If
                End If
            End If
        End If
    End If

    IsSplitArchiveVolume = False
End Function

Private Function BaseNameUpper(ByVal fileName As String) As String
    Dim n As String
    Dim slash As Long

    n = Replace(Trim$(fileName), "/", "\")
    slash = InStrRev(n, "\")
    If slash > 0 Then
        BaseNameUpper = UCase$(Mid$(n, slash + 1))
    Else
        BaseNameUpper = UCase$(n)
    End If
End Function

Private Function IsDigitsOnly(ByVal text As String) As Boolean
    Dim i As Long
    Dim ch As String

    If Len(text) = 0 Then
        IsDigitsOnly = False
        Exit Function
    End If
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If ch < "0" Or ch > "9" Then
            IsDigitsOnly = False
            Exit Function
        End If
    Next i
    IsDigitsOnly = True
End Function

Public Function FileNameOnly(ByVal fullPath As String) As String
    Dim p As String
    Dim slash As Long
    p = Replace(Trim$(fullPath), "/", "\")
    If Len(p) = 0 Then
        FileNameOnly = vbNullString
        Exit Function
    End If
    If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    slash = InStrRev(p, "\")
    If slash = 0 Then
        FileNameOnly = p
    Else
        FileNameOnly = Mid$(p, slash + 1)
    End If
End Function

' Keyword match against a haystack (full path for files, leaf name for folders).
Public Function TextMatchesCriteria(ByVal text As String, ByVal term1 As String, _
                                   ByVal op As String, ByVal term2 As String, _
                                   ByVal flexible As Boolean, _
                                   Optional ByVal extraAndTerm As String = "") As Boolean
    Dim hay As String
    Dim t1 As String
    Dim t2 As String
    Dim extra As String
    Dim hasTerm2 As Boolean
    Dim has1 As Boolean
    Dim has2 As Boolean
    Dim matched As Boolean
    Dim cmp As VbCompareMethod

    t1 = Trim$(term1)
    t2 = Trim$(term2)
    extra = Trim$(extraAndTerm)
    op = UCase$(Trim$(op))
    hasTerm2 = (Len(t2) > 0 And Len(op) > 0)

    If Len(t1) = 0 Then
        TextMatchesCriteria = False
        Exit Function
    End If

    If flexible Then
        hay = NormalizeForMatch(text)
        t1 = NormalizeForMatch(t1)
        If hasTerm2 Then t2 = NormalizeForMatch(t2)
        If Len(extra) > 0 Then extra = NormalizeForMatch(extra)
        If Len(t1) = 0 Then
            TextMatchesCriteria = False
            Exit Function
        End If
        cmp = vbBinaryCompare
    Else
        hay = text
        cmp = vbTextCompare
    End If

    has1 = (InStr(1, hay, t1, cmp) > 0)
    If hasTerm2 Then
        has2 = (InStr(1, hay, t2, cmp) > 0)
        Select Case op
            Case "AND": matched = has1 And has2
            Case "OR": matched = has1 Or has2
            Case "NOT": matched = has1 And (Not has2)
            Case Else: matched = has1
        End Select
    Else
        matched = has1
    End If

    If matched And Len(extra) > 0 Then
        matched = (InStr(1, hay, extra, cmp) > 0)
    End If

    TextMatchesCriteria = matched
End Function

' Folder rows match the folder's own name only — not ancestor names in the path.
Public Function FolderLeafMatchesCriteria(ByVal folderPath As String, ByVal term1 As String, _
                                         ByVal op As String, ByVal term2 As String, _
                                         ByVal flexible As Boolean, _
                                         Optional ByVal extraAndTerm As String = "") As Boolean
    FolderLeafMatchesCriteria = TextMatchesCriteria(FileNameOnly(folderPath), term1, op, term2, _
                                                    flexible, extraAndTerm)
End Function

' Flexible match: lowercase and strip punctuation separators; keep \ and /.
' Single-pass (no per-separator Replace) for search hot path.
Public Function NormalizeForMatch(ByVal text As String) As String
    Dim i As Long
    Dim n As Long
    Dim ch As Long
    Dim buf As String
    Dim outLen As Long
    Dim src As String

    src = LCase$(text)
    n = Len(src)
    If n = 0 Then
        NormalizeForMatch = vbNullString
        Exit Function
    End If

    buf = Space$(n)
    outLen = 0
    For i = 1 To n
        ch = AscW(Mid$(src, i, 1))
        ' Must stay in sync with SEPARATORS_TO_STRIP
        Select Case ch
            Case 32, 33, 35, 36, 37, 38, 39, 40, 41, 43, 44, 45, 46, 59, _
                 61, 64, 91, 93, 94, 95, 96, 123, 124, 125, 126
                ' skip:  .,;-_'()[]{}`~!@#$%^&+=|
            Case Else
                outLen = outLen + 1
                Mid$(buf, outLen, 1) = ChrW$(ch)
        End Select
    Next i

    If outLen = 0 Then
        NormalizeForMatch = vbNullString
    Else
        NormalizeForMatch = Left$(buf, outLen)
    End If
End Function

' needleNorm: optional pre-normalized needle (flexible mode). Empty = normalize inside.
Public Function PathContains(ByVal haystack As String, ByVal needle As String, ByVal flexible As Boolean, _
                             Optional ByVal needleAlreadyNorm As String = "") As Boolean
    Dim n As String
    Dim h As String

    If flexible Then
        If Len(needleAlreadyNorm) > 0 Then
            n = needleAlreadyNorm
        Else
            n = NormalizeForMatch(Trim$(needle))
        End If
        If Len(n) = 0 Then
            PathContains = False
            Exit Function
        End If
        h = NormalizeForMatch(haystack)
        PathContains = (InStr(1, h, n, vbBinaryCompare) > 0)
    Else
        n = Trim$(needle)
        If Len(n) = 0 Then
            PathContains = False
            Exit Function
        End If
        PathContains = (InStr(1, haystack, n, vbTextCompare) > 0)
    End If
End Function

' Resolve mapped drive roots to UNC when possible. Leaves UNC and local paths unchanged.
Public Function ToUncPath(ByVal path As String, Optional ByVal uncRootOverride As String = "") As String
    Dim trimmed As String
    Dim drive As String
    Dim rest As String
    Dim remote As String

    trimmed = Trim$(path)
    If Len(trimmed) = 0 Then
        ToUncPath = vbNullString
        Exit Function
    End If

    ' Already UNC
    If Left$(trimmed, 2) = "\\" Then
        ToUncPath = Replace(trimmed, "/", "\")
        Exit Function
    End If

    ' Mapped drive like I:\folder\...
    If Len(trimmed) >= 3 And Mid$(trimmed, 2, 1) = ":" And (Mid$(trimmed, 3, 1) = "\" Or Mid$(trimmed, 3, 1) = "/") Then
        drive = UCase$(Left$(trimmed, 2))
        rest = Mid$(trimmed, 3)
        rest = Replace(rest, "/", "\")
        If Left$(rest, 1) = "\" Then rest = Mid$(rest, 2)

        remote = GetRemoteNameForDrive(drive)
        If Len(remote) > 0 Then
            ToUncPath = JoinUnc(remote, rest)
            Exit Function
        End If

        If Len(Trim$(uncRootOverride)) > 0 Then
            ToUncPath = JoinUnc(Trim$(uncRootOverride), rest)
            Exit Function
        End If
    End If

    ToUncPath = Replace(trimmed, "/", "\")
End Function

' Display without drive letter or \\server\share\ prefix.
' Relative part is returned with a leading "\" when non-empty (e.g. "\Folder\Sub").
Public Function ToDisplayPath(ByVal fullPath As String) As String
    Dim p As String
    Dim slashPos As Long
    Dim second As Long
    Dim rel As String

    p = Trim$(fullPath)
    p = Replace(p, "/", "\")

    If Left$(p, 2) = "\\" Then
        ' \\server\share\rest -> \rest
        slashPos = InStr(3, p, "\")
        If slashPos > 0 Then
            second = InStr(slashPos + 1, p, "\")
            If second > 0 Then
                rel = Mid$(p, second + 1)
            Else
                rel = vbNullString
            End If
        Else
            rel = vbNullString
        End If
    ElseIf Len(p) >= 3 And Mid$(p, 2, 1) = ":" And Mid$(p, 3, 1) = "\" Then
        ' I:\mssu\... -> \mssu\...
        rel = Mid$(p, 4)
    Else
        ToDisplayPath = p
        Exit Function
    End If

    If Len(rel) = 0 Then
        ToDisplayPath = "\"
    ElseIf Left$(rel, 1) = "\" Then
        ToDisplayPath = rel
    Else
        ToDisplayPath = "\" & rel
    End If
End Function

' \\server\share\rest -> \\server\share ; I:\rest -> I: (or UNC share if mapped)
Public Function UncShareRoot(ByVal fullPath As String) As String
    Dim p As String
    Dim slashPos As Long
    Dim second As Long
    Dim unc As String

    p = Trim$(Replace(fullPath, "/", "\"))
    If Len(p) = 0 Then
        UncShareRoot = vbNullString
        Exit Function
    End If

    If Left$(p, 2) = "\\" Then
        slashPos = InStr(3, p, "\")
        If slashPos <= 0 Then
            UncShareRoot = p
            Exit Function
        End If
        second = InStr(slashPos + 1, p, "\")
        If second > 0 Then
            UncShareRoot = Left$(p, second - 1)
        Else
            UncShareRoot = p
        End If
        Exit Function
    End If

    If Len(p) >= 2 And Mid$(p, 2, 1) = ":" Then
        unc = ToUncPath(Left$(p, 2) & "\", "")
        If Left$(unc, 2) = "\\" Then
            UncShareRoot = UncShareRoot(unc)
        Else
            UncShareRoot = UCase$(Left$(p, 2))
        End If
        Exit Function
    End If

    UncShareRoot = vbNullString
End Function

' Friendly label for a UNC share, e.g. "SHARE_Public" from mapped "SHARE_Public (K:)".
' Falls back to share leaf (without trailing $) when not mapped / no volume name.
Public Function FriendlyDriveLabel(ByVal fullPathOrShare As String) As String
    Dim share As String
    Dim fso As Object
    Dim d As Object
    Dim remote As String
    Dim vol As String
    Dim shellName As String
    Dim leaf As String
    Dim driveUnc As String
    Dim letter As String

    share = UncShareRoot(fullPathOrShare)
    If Len(share) = 0 Then
        FriendlyDriveLabel = vbNullString
        Exit Function
    End If

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then
        For Each d In fso.Drives
            Err.Clear
            If d.DriveType = 3 Then ' Network
                If d.IsReady Then
                    letter = UCase$(Trim$(CStr(d.DriveLetter & "")))
                    remote = CStr(d.ShareName & "")
                    If Len(remote) > 0 Then
                        If Right$(remote, 1) = "\" Then remote = Left$(remote, Len(remote) - 1)
                    End If
                    ' ShareName is often blank; resolve via WNet like ToUncPath
                    If Len(remote) = 0 And Len(letter) = 1 Then
                        remote = GetRemoteNameForDrive(letter & ":")
                        If Right$(remote, 1) = "\" Then remote = Left$(remote, Len(remote) - 1)
                    End If
                    If Len(remote) = 0 And Len(letter) = 1 Then
                        driveUnc = ToUncPath(letter & ":\", "")
                        remote = UncShareRoot(driveUnc)
                    End If

                    If Len(remote) > 0 And StrComp(remote, share, vbTextCompare) = 0 Then
                        ' Prefer Explorer-style name: "SHARE_Public (K:)" -> SHARE_Public
                        If Len(letter) = 1 Then
                            shellName = ShellDriveDisplayName(letter & ":")
                            If Len(shellName) > 0 Then
                                FriendlyDriveLabel = StripDriveLetterSuffix(shellName)
                                If Len(FriendlyDriveLabel) > 0 Then
                                    On Error GoTo 0
                                    Exit Function
                                End If
                            End If
                        End If
                        vol = Trim$(CStr(d.VolumeName & ""))
                        If Len(vol) > 0 Then
                            FriendlyDriveLabel = vol
                            On Error GoTo 0
                            Exit Function
                        End If
                    End If
                End If
            End If
        Next d
    End If
    On Error GoTo 0

    ' Fallback: last segment of \\server\share$ -> share (drop trailing $)
    leaf = Mid$(share, InStrRev(share, "\") + 1)
    If Right$(leaf, 1) = "$" Then leaf = Left$(leaf, Len(leaf) - 1)
    FriendlyDriveLabel = leaf
End Function

Private Function ShellDriveDisplayName(ByVal driveWithColon As String) As String
    Dim sh As Object
    Dim ns As Object
    Dim it As Object
    On Error Resume Next
    Set sh = CreateObject("Shell.Application")
    If sh Is Nothing Then Exit Function
    Set ns = sh.Namespace(driveWithColon & "\")
    If ns Is Nothing Then Exit Function
    Set it = ns.Self
    If it Is Nothing Then Exit Function
    ShellDriveDisplayName = CStr(it.Name & "")
    On Error GoTo 0
End Function

' "SHARE_Public (K:)" -> "SHARE_Public"
Private Function StripDriveLetterSuffix(ByVal displayName As String) As String
    Dim p As Long
    Dim n As String
    n = Trim$(displayName)
    p = InStrRev(n, " (")
    If p > 1 And Right$(n, 1) = ")" Then
        StripDriveLetterSuffix = Trim$(Left$(n, p - 1))
    Else
        StripDriveLetterSuffix = n
    End If
End Function

Public Function EnsureTrailingSlash(ByVal path As String) As String
    If Len(path) = 0 Then
        EnsureTrailingSlash = vbNullString
    ElseIf Right$(path, 1) = "\" Then
        EnsureTrailingSlash = path
    Else
        EnsureTrailingSlash = path & "\"
    End If
End Function

Public Function PathStartsWithRoot(ByVal filePath As String, ByVal rootPath As String) As Boolean
    Dim f As String
    Dim r As String
    f = LCase$(Replace(filePath, "/", "\"))
    r = LCase$(EnsureTrailingSlash(Replace(rootPath, "/", "\")))
    If Len(r) = 0 Then
        PathStartsWithRoot = False
        Exit Function
    End If
    PathStartsWithRoot = (Left$(f, Len(r)) = r) Or (f = Left$(r, Len(r) - 1))
End Function

Private Function JoinUnc(ByVal uncRoot As String, ByVal relativePath As String) As String
    Dim root As String
    Dim rel As String
    root = RTrim$(Replace(uncRoot, "/", "\"))
    Do While Right$(root, 1) = "\"
        root = Left$(root, Len(root) - 1)
    Loop
    rel = Replace(relativePath, "/", "\")
    Do While Left$(rel, 1) = "\"
        rel = Mid$(rel, 2)
    Loop
    If Len(rel) = 0 Then
        JoinUnc = root
    Else
        JoinUnc = root & "\" & rel
    End If
End Function

Private Function GetRemoteNameForDrive(ByVal driveWithColon As String) As String
    Dim buf As String
    Dim bufSize As Long
    Dim rc As Long
    Dim nullPos As Long

    bufSize = 1024
    buf = String$(bufSize, vbNullChar)
    rc = WNetGetConnection(driveWithColon, buf, bufSize)
    If rc = 0 Then
        nullPos = InStr(1, buf, vbNullChar)
        If nullPos > 0 Then
            GetRemoteNameForDrive = Left$(buf, nullPos - 1)
        Else
            GetRemoteNameForDrive = buf
        End If
    Else
        GetRemoteNameForDrive = vbNullString
    End If
End Function

' Folder picker (Explorer-style). Returns empty string if cancelled.
Public Function BrowseForFolder(Optional ByVal dialogTitle As String = "Select a folder to ingest") As String
    Dim dlg As FileDialog
    Dim picked As String

    Set dlg = Application.FileDialog(msoFileDialogFolderPicker)
    With dlg
        .Title = dialogTitle
        .AllowMultiSelect = False
        If .Show <> -1 Then
            BrowseForFolder = vbNullString
            Exit Function
        End If
        picked = .SelectedItems(1)
    End With

    BrowseForFolder = picked
End Function
