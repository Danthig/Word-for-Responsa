Attribute VB_Name = "BarIlanSearch"
Option Explicit
'==============================================================================
' Search the text selected in Word inside Bar-Ilan Responsa (RESPONSA.exe)
'   SearchSelectionInBarIlan          - plain words, in the "easy search" window
'   SearchSelectionAdvancedInBarIlan  - "advanced search" window: every word wrapped with
'                                       prefix/suffix, with a distance in front, e.g.
'                                       30: #!word1# #!word2# #!word3#
'   SearchSelectionInBarIlanCitation - search the selected text in "Browse > Write Sources"
' Non-contiguous (Ctrl) selections are supported: all selected areas are searched.
' Window IDs / commands taken from helper/lib/src/native/responsa_search_automation.dart
' Word 2010+ (VBA7). Hebrew strings are built with ChrW so the file stays ASCII.
'==============================================================================

'---- settings ----------------------------------------------------------------
Private Const CMD_OPEN_SEARCH As Long = 32857
Private Const CMD_OPEN_CITATION As Long = 32781
Private Const ID_SEARCH_EDIT As Long = 1233
Private Const ID_CITATION_EDIT As Long = 1021
Private Const ID_CITATION_SEARCH As Long = 1187
Private Const ID_CITATION_TAB As Long = 12320
Private Const ID_EASY_BTN As Long = 1207       ' mode button observed in Responsa v29
Private Const ID_ADVANCED_BTN As Long = 1209    ' button that switches to "advanced search"
Private Const ID_ADVANCED_ONLY As Long = 1065   ' control that exists only in advanced mode
Private Const CITATION_TAB_INDEX As Long = 1    ' "Write Sources" tab in Responsa v29
Private Const MAX_WORDS As Long = 10
Private Const MAX_EDITS As Long = 4
Private Const DIALOG_WAIT_SEC As Double = 12
Private Const RESEND_EVERY_SEC As Double = 3
Private Const FALLBACK_AFTER_SEC As Double = 3  ' no title match -> accept any search dialog
Private Const FOCUS_STRONG As Boolean = False   ' True = also SwitchToThisWindow (if focus fails)

' ---- advanced search syntax (edit here) ----
Private Const ADV_PREFIX As String = "#!"       ' placed before every word
Private Const ADV_SUFFIX As String = "#"        ' placed after every word
Private Const ADV_DISTANCE As String = "30:"    ' distance, written once in front of the words

'---- Win32 -------------------------------------------------------------------
Private Const WM_COMMAND As Long = &H111
Private Const WM_SETTEXT As Long = &HC
Private Const WM_GETTEXT As Long = &HD
Private Const EN_CHANGE As Long = &H300
Private Const BM_CLICK As Long = &HF5
Private Const TCM_SETCURFOCUS As Long = &H1330
Private Const SMTO_ABORTIFHUNG As Long = 2
Private Const GW_OWNER As Long = 4
Private Const SW_RESTORE As Long = 9
Private Const SW_SHOW As Long = 5
Private Const PROCESS_QUERY_LIMITED_INFORMATION As Long = &H1000
Private Const CF_UNICODETEXT As Long = 13
Private Const GMEM_MOVEABLE As Long = 2

Private Declare PtrSafe Function EnumWindows Lib "user32" (ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function EnumChildWindows Lib "user32" (ByVal hWndParent As LongPtr, ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As LongPtr, ByRef lpdwProcessId As Long) As Long
Private Declare PtrSafe Function GetParent Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function IsIconic Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function SetFocus Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function BringWindowToTop Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Sub SwitchToThisWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal fAltTab As Long)
Private Declare PtrSafe Function GetLastActivePopup Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal uCmd As Long) As LongPtr
Private Declare PtrSafe Function GetMenu Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetDlgItem Lib "user32" (ByVal hDlg As LongPtr, ByVal nIDDlgItem As Long) As LongPtr
Private Declare PtrSafe Function GetDlgCtrlID Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function GetClassNameW Lib "user32" (ByVal hWnd As LongPtr, ByVal lpClassName As LongPtr, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function PostMessageW Lib "user32" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function SendMessageTimeoutW Lib "user32" (ByVal hWnd As LongPtr, ByVal Msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr, ByVal fuFlags As Long, ByVal uTimeout As Long, ByRef lpdwResult As LongPtr) As LongPtr
Private Declare PtrSafe Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare PtrSafe Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As LongPtr
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function QueryFullProcessImageNameW Lib "kernel32" (ByVal hProcess As LongPtr, ByVal dwFlags As Long, ByVal lpExeName As LongPtr, ByRef lpdwSize As Long) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWndNewOwner As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal uFormat As Long) As LongPtr
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal uFormat As Long, ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal uFormat As Long) As Long
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As LongPtr)

'---- module state (enumeration callbacks) ------------------------------------
Private mWins As Collection
Private mLastPid As Long
Private mLastIsResp As Boolean
Private mTarget As String
Private mBtn As LongPtr
Private mEdits As Long
Private mFindId As Long
Private mFoundId As LongPtr

'==============================================================================
' Entry points
'==============================================================================
Public Sub SearchSelectionInBarIlan()
    RunSearch False
End Sub

Public Sub SearchSelectionAdvancedInBarIlan()
    RunSearch True
End Sub

Public Sub SearchSelectionInBarIlanCitation()
    RunCitationSearch
End Sub

Private Sub RunSearch(ByVal advanced As Boolean)
    Dim q As String, promptFound As Boolean
    If Selection.Type = wdSelectionIP Then
        MsgBox "No text selected.", vbInformation
        Exit Sub
    End If

    q = CleanQuery(GetSelectedText())
    If Len(q) = 0 Then
        MsgBox "No Hebrew words found in the selection.", vbInformation
        Exit Sub
    End If
    If advanced Then q = BuildAdvancedQuery(q)

    Dim hMain As LongPtr
    hMain = EnsureResponsaRunning()
    If hMain = 0 Then
        MsgBox "Could not find or start RESPONSA. Start it manually or set the RESPONSA_PATH environment variable to the full path of RESPONSA.exe.", vbExclamation
        Exit Sub
    End If

    CloseLeftoverModals
    hMain = FindMainWindow()
    If hMain = 0 Then Exit Sub

    Dim hDlg As LongPtr, hEdit As LongPtr, hBtn As LongPtr
    hDlg = EnsureSearchDialog(hMain, advanced, hBtn, hEdit)
    If hDlg = 0 Then
        MsgBox "The RESPONSA search window was not found.", vbExclamation
        Exit Sub
    End If

    If Not SetText(hEdit, q) Then
        MsgBox "Could not write the text into the search box.", vbExclamation
        Exit Sub
    End If
    If PostMessageW(hBtn, BM_CLICK, 0, 0) = 0 Then
        MsgBox "Could not start the RESPONSA search.", vbExclamation
        Exit Sub
    End If
    If Not FocusSearchAllPrompt(hMain, promptFound) Then
        If Not promptFound Then BringToFront hMain
    End If
End Sub

Private Sub RunCitationSearch()
    Dim q As String
    If Selection.Type = wdSelectionIP Then
        MsgBox "No text selected.", vbInformation
        Exit Sub
    End If

    q = NormalizeCitationText(GetSelectedText())
    If Len(q) = 0 Then
        MsgBox "The selected text is empty.", vbInformation
        Exit Sub
    End If

    Dim hMain As LongPtr
    hMain = EnsureResponsaRunning()
    If hMain = 0 Then
        MsgBox "Could not find or start RESPONSA. Start it manually or set the RESPONSA_PATH environment variable to the full path of RESPONSA.exe.", vbExclamation
        Exit Sub
    End If

    CloseLeftoverModals
    hMain = FindMainWindow()
    If hMain = 0 Then
        MsgBox "The RESPONSA window was not found.", vbExclamation
        Exit Sub
    End If

    Dim hDlg As LongPtr, hEdit As LongPtr, hBtn As LongPtr
    hDlg = EnsureCitationDialog(hMain, hEdit, hBtn)
    If hDlg = 0 Then
        MsgBox "The RESPONSA 'Write Sources' window was not found.", vbExclamation
        Exit Sub
    End If

    If Not SetWindowTextOnly(hEdit, q) Then
        MsgBox "Could not write the selected text into 'Write Sources'.", vbExclamation
        Exit Sub
    End If
    If PostMessageW(hBtn, BM_CLICK, 0, 0) = 0 Then
        MsgBox "Could not start the source search.", vbExclamation
        Exit Sub
    End If
    Sleep 300
    If Not BringCitationToFront(hMain, hDlg, hEdit) Then
        MsgBox "The source search started, but the 'Write Sources' window could not be returned to the foreground.", vbExclamation
    End If
End Sub

'==============================================================================
' Query building
'==============================================================================
' Hebrew words only, no niqqud/cantillation/operators, max MAX_WORDS
Private Function CleanQuery(ByVal s As String) As String
    Dim i As Long, c As Long, buf As String
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1))
        If c < 0 Then c = c + 65536
        Select Case c
            Case 1488 To 1514
                buf = buf & ChrW(c)
            Case 1425 To 1479
                If c = 1470 Then buf = buf & " "
            Case 34, 8220, 8221, 1524
                buf = buf & """"
            Case 39, 8216, 8217, 1523
                buf = buf & "'"
            Case Else
                buf = buf & " "
        End Select
    Next

    Dim parts() As String, w As String, n As Long, out As String
    parts = Split(buf, " ")
    For i = LBound(parts) To UBound(parts)
        w = parts(i)
        Do While Len(w) > 0 And (Left$(w, 1) = """" Or Left$(w, 1) = "'")
            w = Mid$(w, 2)
        Loop
        Do While Len(w) > 0 And (Right$(w, 1) = """" Or Right$(w, 1) = "'")
            w = Left$(w, Len(w) - 1)
        Loop
        If Len(w) > 0 Then
            n = n + 1
            If n > MAX_WORDS Then Exit For
            If Len(out) > 0 Then out = out & " "
            out = out & w
        End If
    Next
    CleanQuery = out
End Function

' "w1 w2 w3"  ->  "30: #!w1# #!w2# #!w3#"   (a single word gets no distance)
Private Function BuildAdvancedQuery(ByVal q As String) As String
    Dim parts() As String, i As Long, out As String
    parts = Split(q, " ")
    For i = LBound(parts) To UBound(parts)
        If Len(out) > 0 Then out = out & " "
        out = out & ADV_PREFIX & parts(i) & ADV_SUFFIX
    Next
    If UBound(parts) > LBound(parts) Then out = ADV_DISTANCE & " " & out
    BuildAdvancedQuery = out
End Function

Private Function NormalizeCitationText(ByVal s As String) As String
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, ChrW(7), " ")
    s = Replace(s, ChrW(11), " ")
    s = Replace(s, ChrW(160), " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormalizeCitationText = Trim$(s)
End Function

'==============================================================================
' Selected text, including non-contiguous (Ctrl) selections.
' Selection.Text returns only one of the areas, so the selection is also copied
' and the clipboard text is read; whichever yields more Hebrew words wins.
' The clipboard text that was there before is restored (text only).
'==============================================================================
Private Function GetSelectedText() As String
    Dim direct As String, clip As String, saved As String
    direct = Selection.Text

    saved = ReadClipboardText()
    WriteClipboardText ChrW(1)                  ' sentinel: tells a failed copy apart
    On Error Resume Next
    Selection.Copy
    On Error GoTo 0
    Sleep 50
    clip = ReadClipboardText()
    If Len(saved) > 0 Then WriteClipboardText saved

    If clip = ChrW(1) Or Len(clip) = 0 Then
        GetSelectedText = direct
    ElseIf CountWords(CleanQuery(clip)) > CountWords(CleanQuery(direct)) Then
        GetSelectedText = clip
    Else
        GetSelectedText = direct
    End If
End Function

Private Function CountWords(ByVal q As String) As Long
    If Len(q) = 0 Then Exit Function
    CountWords = UBound(Split(q, " ")) + 1
End Function

Private Function ReadClipboardText() As String
    Dim h As LongPtr, p As LongPtr, n As Long, buf As String, tries As Long
    If IsClipboardFormatAvailable(CF_UNICODETEXT) = 0 Then Exit Function
    Do While OpenClipboard(0) = 0
        tries = tries + 1
        If tries > 10 Then Exit Function
        Sleep 20
    Loop
    h = GetClipboardData(CF_UNICODETEXT)
    If h <> 0 Then
        p = GlobalLock(h)
        If p <> 0 Then
            n = lstrlenW(p)
            If n > 0 Then
                buf = String$(n, vbNullChar)
                RtlMoveMemory StrPtr(buf), p, n * 2
                ReadClipboardText = buf
            End If
            GlobalUnlock h
        End If
    End If
    CloseClipboard
End Function

Private Sub WriteClipboardText(ByVal s As String)
    Dim h As LongPtr, p As LongPtr, bytes As Long, tries As Long
    bytes = (Len(s) + 1) * 2
    h = GlobalAlloc(GMEM_MOVEABLE, bytes)
    If h = 0 Then Exit Sub
    p = GlobalLock(h)
    If p = 0 Then Exit Sub
    RtlMoveMemory p, StrPtr(s), Len(s) * 2
    RtlMoveMemory p + Len(s) * 2, StrPtr(vbNullChar), 2   ' terminating null
    GlobalUnlock h
    Do While OpenClipboard(0) = 0
        tries = tries + 1
        If tries > 10 Then Exit Sub
        Sleep 20
    Loop
    EmptyClipboard
    SetClipboardData CF_UNICODETEXT, h
    CloseClipboard
End Sub

'==============================================================================
' Locating RESPONSA windows
'==============================================================================
Private Function IsResponsaWindow(ByVal hWnd As LongPtr) As Boolean
    Dim pid As Long, hProc As LongPtr, buf As String, sz As Long
    GetWindowThreadProcessId hWnd, pid
    If pid = 0 Then Exit Function
    If pid = mLastPid Then
        IsResponsaWindow = mLastIsResp
        Exit Function
    End If
    mLastPid = pid
    mLastIsResp = False
    hProc = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, 0, pid)
    If hProc <> 0 Then
        buf = String$(1024, vbNullChar)
        sz = 1024
        If QueryFullProcessImageNameW(hProc, 0, StrPtr(buf), sz) <> 0 Then
            mLastIsResp = (LCase$(Right$(Left$(buf, sz), 13)) = "\responsa.exe")
        End If
        CloseHandle hProc
    End If
    IsResponsaWindow = mLastIsResp
End Function

Private Function EnumTopProc(ByVal hWnd As LongPtr, ByVal lParam As LongPtr) As Long
    If IsResponsaWindow(hWnd) Then mWins.Add hWnd
    EnumTopProc = 1
End Function

Private Sub CollectResponsaWindows()
    Set mWins = New Collection
    mLastPid = 0
    EnumWindows AddressOf EnumTopProc, 0
End Sub

Private Function FindMainWindow() As LongPtr
    Dim v As Variant, h As LongPtr
    CollectResponsaWindows
    For Each v In mWins
        h = CLngPtr(v)
        If IsWindowVisible(h) <> 0 Then
            If GetWindow(h, GW_OWNER) = 0 Then
                If GetMenu(h) <> 0 Then
                    FindMainWindow = h
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Private Function EnsureResponsaRunning() As LongPtr
    Dim h As LongPtr, t0 As Double, responsaPath As String
    h = FindMainWindow()
    If h <> 0 Then
        EnsureResponsaRunning = h
        Exit Function
    End If
    responsaPath = Trim$(Environ$("RESPONSA_PATH"))
    If Len(responsaPath) = 0 Then Exit Function
    If Left$(responsaPath, 1) = """" And Right$(responsaPath, 1) = """" Then
        responsaPath = Mid$(responsaPath, 2, Len(responsaPath) - 2)
    End If
    If Len(Dir$(responsaPath)) = 0 Then Exit Function
    Shell """" & responsaPath & """", vbNormalFocus
    t0 = Timer
    Do
        Sleep 500
        DoEvents
        h = FindMainWindow()
        If h <> 0 Then Sleep 1500: Exit Do
    Loop While Timer - t0 < 40
    EnsureResponsaRunning = h
End Function

'==============================================================================
' Controls / text helpers
'==============================================================================
Private Function ClassOf(ByVal h As LongPtr) As String
    Dim buf As String, n As Long
    buf = String$(256, vbNullChar)
    n = GetClassNameW(h, StrPtr(buf), 256)
    ClassOf = Left$(buf, n)
End Function

Private Function TextOf(ByVal h As LongPtr) As String
    Dim buf As String, r As LongPtr
    buf = String$(512, vbNullChar)
    If SendMessageTimeoutW(h, WM_GETTEXT, 512, StrPtr(buf), SMTO_ABORTIFHUNG, 300, r) = 0 Then Exit Function
    TextOf = Left$(buf, CLng(r))
End Function

Private Function SetText(ByVal h As LongPtr, ByVal s As String) As Boolean
    Dim r As LongPtr, hParent As LongPtr, notify As LongPtr
    If SendMessageTimeoutW(h, WM_SETTEXT, 0, StrPtr(s), SMTO_ABORTIFHUNG, 2000, r) = 0 Then Exit Function
    If TextOf(h) <> s Then Exit Function

    ' WM_SETTEXT changes the edit's text but does not send the EN_CHANGE
    ' notification that updates Responsa's remembered search query.
    hParent = GetParent(h)
    If hParent = 0 Then Exit Function
    notify = CLngPtr((GetDlgCtrlID(h) And &HFFFF&) Or (EN_CHANGE * &H10000))
    SetText = (SendMessageTimeoutW(hParent, WM_COMMAND, notify, h, SMTO_ABORTIFHUNG, 2000, r) <> 0)
End Function

Private Function SetWindowTextOnly(ByVal h As LongPtr, ByVal s As String) As Boolean
    Dim r As LongPtr
    If SendMessageTimeoutW(h, WM_SETTEXT, 0, StrPtr(s), SMTO_ABORTIFHUNG, 2000, r) = 0 Then Exit Function
    SetWindowTextOnly = (TextOf(h) = s)
End Function

Private Function NormCaption(ByVal s As String) As String
    s = Replace(s, "&", "")
    s = Replace(s, ChrW(8206), "")
    s = Replace(s, ChrW(8207), "")
    NormCaption = Trim$(s)
End Function

Private Function EnumChildProc(ByVal hWnd As LongPtr, ByVal lParam As LongPtr) As Long
    Dim cls As String
    cls = LCase$(ClassOf(hWnd))
    If cls = "button" Then
        If mBtn = 0 And Len(mTarget) > 0 Then
            If NormCaption(TextOf(hWnd)) = mTarget Then mBtn = hWnd
        End If
    ElseIf cls = "edit" Then
        mEdits = mEdits + 1
    End If
    EnumChildProc = 1
End Function

Private Function FindButton(ByVal hParent As LongPtr, ByVal caption As String) As LongPtr
    mTarget = caption
    mBtn = 0
    mEdits = 0
    EnumChildWindows hParent, AddressOf EnumChildProc, 0
    FindButton = mBtn
End Function

Private Function EnumIdProc(ByVal hWnd As LongPtr, ByVal lParam As LongPtr) As Long
    If GetDlgCtrlID(hWnd) = mFindId Then
        mFoundId = hWnd
        EnumIdProc = 0
    Else
        EnumIdProc = 1
    End If
End Function

' child control (any depth) by control ID; 0 if none
Private Function ChildById(ByVal hParent As LongPtr, ByVal id As Long) As LongPtr
    mFindId = id
    mFoundId = 0
    EnumChildWindows hParent, AddressOf EnumIdProc, 0
    ChildById = mFoundId
End Function

'==============================================================================
' Search dialog. RESPONSA creates all search dialogs (easy / table / advanced /
' free text) and hides all but the selected one. Only the visible dialog is safe
' to use: a hidden dialog can display a new query without updating the active
' search state used by the "search all databases" prompt.
'==============================================================================
Private Function FindSearchDialog(ByVal mode As Long, ByRef btn As LongPtr, ByRef edt As LongPtr) As LongPtr
    Dim v As Variant, h As LongPtr, b As LongPtr, e As LongPtr
    Dim title As String, isAdv As Boolean, ok As Boolean
    CollectResponsaWindows
    For Each v In mWins
        h = CLngPtr(v)
        If IsWindowVisible(h) <> 0 Then
            e = GetDlgItem(h, ID_SEARCH_EDIT)
            If e <> 0 Then
                b = FindButton(h, TxtSearch())
                If b <> 0 And mEdits <= MAX_EDITS Then
                    title = NormCaption(TextOf(h))
                    isAdv = (ChildById(h, ID_ADVANCED_ONLY) <> 0)
                    Select Case mode
                        Case 0: ok = (Not isAdv) And (InStr(title, TxtEasyTitle()) > 0)
                        Case 1: ok = isAdv Or (InStr(title, TxtAdvTitle()) > 0)
                        Case Else: ok = True
                    End Select
                    If ok Then
                        btn = b
                        edt = e
                        FindSearchDialog = h
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
End Function

' open the search dialogs (command resent every few seconds) and return the one
' for the wanted mode. If another search mode is visible, switch that dialog
' instead of filling one of Responsa's hidden, inactive dialogs.
Private Function EnsureSearchDialog(ByVal hMain As LongPtr, ByVal advanced As Boolean, ByRef btn As LongPtr, ByRef edt As LongPtr) As LongPtr
    Dim h As LongPtr, t0 As Double, tSent As Double, mode As Long, switchAttempted As Boolean
    mode = IIf(advanced, 1, 0)

    h = FindSearchDialog(mode, btn, edt)
    If h <> 0 Then EnsureSearchDialog = h: Exit Function

    h = FindSearchDialog(2, btn, edt)
    If h <> 0 Then
        switchAttempted = True
        If advanced Then
            EnsureSearchDialog = EnsureAdvancedMode(h, btn, edt)
        Else
            EnsureSearchDialog = EnsureEasyMode(h, btn, edt)
        End If
        If EnsureSearchDialog <> 0 Then Exit Function
    End If

    t0 = Timer
    tSent = -100
    Do
        If Timer - tSent >= RESEND_EVERY_SEC Then
            PostMessageW hMain, WM_COMMAND, CMD_OPEN_SEARCH, 0
            tSent = Timer
        End If
        Sleep 150
        DoEvents
        h = FindSearchDialog(mode, btn, edt)
        If h <> 0 Then EnsureSearchDialog = h: Exit Function
        If Not switchAttempted And Timer - t0 >= FALLBACK_AFTER_SEC Then
            h = FindSearchDialog(2, btn, edt)
            If h <> 0 Then
                switchAttempted = True
                If advanced Then
                    EnsureSearchDialog = EnsureAdvancedMode(h, btn, edt)
                Else
                    EnsureSearchDialog = EnsureEasyMode(h, btn, edt)
                End If
                If EnsureSearchDialog <> 0 Then Exit Function
            End If
        End If
    Loop While Timer - t0 < DIALOG_WAIT_SEC
End Function

' switch the visible dialog to "advanced search" with its mode button
Private Function EnsureAdvancedMode(ByVal hDlg As LongPtr, ByRef btn As LongPtr, ByRef edt As LongPtr) As LongPtr
    Dim b As LongPtr, d As LongPtr, t0 As Double
    If ChildById(hDlg, ID_ADVANCED_ONLY) <> 0 Or InStr(NormCaption(TextOf(hDlg)), TxtAdvTitle()) > 0 Then
        EnsureAdvancedMode = FindSearchDialog(1, btn, edt)
        Exit Function
    End If

    b = ChildById(hDlg, ID_ADVANCED_BTN)
    If b = 0 Then Exit Function
    If PostMessageW(b, BM_CLICK, 0, 0) = 0 Then Exit Function

    t0 = Timer
    Do
        Sleep 150
        DoEvents
        d = FindSearchDialog(1, btn, edt)
        If d <> 0 Then EnsureAdvancedMode = d: Exit Function
    Loop While Timer - t0 < 6
End Function

Private Function EnsureEasyMode(ByVal hDlg As LongPtr, ByRef btn As LongPtr, ByRef edt As LongPtr) As LongPtr
    Dim b As LongPtr, d As LongPtr, t0 As Double
    If InStr(NormCaption(TextOf(hDlg)), TxtEasyTitle()) > 0 Then
        EnsureEasyMode = FindSearchDialog(0, btn, edt)
        Exit Function
    End If

    b = ChildById(hDlg, ID_EASY_BTN)
    If b = 0 Then Exit Function
    If PostMessageW(b, BM_CLICK, 0, 0) = 0 Then Exit Function

    t0 = Timer
    Do
        Sleep 150
        DoEvents
        d = FindSearchDialog(0, btn, edt)
        If d <> 0 Then EnsureEasyMode = d: Exit Function
    Loop While Timer - t0 < 6
End Function

'==============================================================================
' Citation search dialog ("Browse > Write Sources")
'==============================================================================
Private Function FindCitationDialog(ByVal visibleOnly As Boolean, ByRef edt As LongPtr, ByRef btn As LongPtr) As LongPtr
    Dim v As Variant, h As LongPtr, title As String
    CollectResponsaWindows
    For Each v In mWins
        h = CLngPtr(v)
        If (Not visibleOnly) Or IsWindowVisible(h) <> 0 Then
            title = NormCaption(TextOf(h))
            If InStr(1, title, TxtBrowseTitle(), vbTextCompare) > 0 Then
                edt = ChildById(h, ID_CITATION_EDIT)
                btn = ChildById(h, ID_CITATION_SEARCH)
                If edt <> 0 And btn <> 0 Then
                    FindCitationDialog = h
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Private Function FocusCitationTab(ByVal hDlg As LongPtr, ByRef edt As LongPtr, ByRef btn As LongPtr) As Boolean
    Dim hTab As LongPtr, r As LongPtr
    hTab = ChildById(hDlg, ID_CITATION_TAB)
    If hTab = 0 Then Exit Function

    If SendMessageTimeoutW(hTab, TCM_SETCURFOCUS, CITATION_TAB_INDEX, 0, SMTO_ABORTIFHUNG, 2000, r) = 0 Then Exit Function
    Sleep 800

    edt = ChildById(hDlg, ID_CITATION_EDIT)
    btn = ChildById(hDlg, ID_CITATION_SEARCH)
    FocusCitationTab = (edt <> 0 And btn <> 0)
End Function

Private Function EnsureCitationDialog(ByVal hMain As LongPtr, ByRef edt As LongPtr, ByRef btn As LongPtr) As LongPtr
    Dim h As LongPtr, t0 As Double, tSent As Double, needsShow As Boolean
    h = FindCitationDialog(False, edt, btn)
    If h = 0 Then
        t0 = Timer
        tSent = -100
        Do
            If Timer - tSent >= RESEND_EVERY_SEC Then
                PostMessageW hMain, WM_COMMAND, CMD_OPEN_CITATION, 0
                tSent = Timer
            End If
            Sleep 150
            DoEvents
            h = FindCitationDialog(False, edt, btn)
        Loop While h = 0 And Timer - t0 < DIALOG_WAIT_SEC
    End If
    If h = 0 Then Exit Function

    If Not FocusCitationTab(h, edt, btn) Then Exit Function
    needsShow = (IsWindowVisible(h) = 0)
    If needsShow Then
        t0 = Timer
        tSent = -100
    End If

    t0 = Timer
    Do
        If needsShow And Timer - tSent >= RESEND_EVERY_SEC Then
            PostMessageW hMain, WM_COMMAND, CMD_OPEN_CITATION, 0
            tSent = Timer
        End If

        h = FindCitationDialog(True, edt, btn)
        If h = 0 Then
            needsShow = True
        Else
            needsShow = False
            If FocusCitationTab(h, edt, btn) Then
                If IsWindowVisible(edt) <> 0 And IsWindowVisible(btn) <> 0 Then
                    BringToFront hMain
                    EnsureCitationDialog = h
                    Exit Function
                End If
            End If
        End If

        Sleep 150
        DoEvents
    Loop While Timer - t0 < DIALOG_WAIT_SEC

    If h <> 0 And IsWindowVisible(h) <> 0 Then
        If FocusCitationTab(h, edt, btn) And IsWindowVisible(edt) <> 0 And IsWindowVisible(btn) <> 0 Then
            BringToFront hMain
            EnsureCitationDialog = h
        End If
    End If
End Function

'==============================================================================
' Modals: leftovers before the search, result modal after it
'==============================================================================
Private Function IsInfoOrResultModal(ByVal h As LongPtr) As Boolean
    Dim title As String
    If IsWindowVisible(h) = 0 Then Exit Function
    If ClassOf(h) <> "#32770" Then Exit Function
    If GetDlgItem(h, ID_SEARCH_EDIT) <> 0 Then Exit Function
    title = TextOf(h)
    IsInfoOrResultModal = (title = TxtInfoTitle() Or InStr(title, TxtResultsWord()) > 0)
End Function

Private Sub CloseLeftoverModals()
    Dim v As Variant, h As LongPtr, b As LongPtr, closed As Boolean
    CollectResponsaWindows
    For Each v In mWins
        h = CLngPtr(v)
        If IsInfoOrResultModal(h) Then
            b = FindButton(h, TxtOk())
            If b = 0 Then b = FindButton(h, TxtCancel())
            If b <> 0 Then PostMessageW b, BM_CLICK, 0, 0: closed = True
        End If
    Next
    If closed Then Sleep 300: DoEvents          ' only wait if something was closed
End Sub

' Wait briefly for the no-results prompt. Do not bring the search window back
' to the foreground while that prompt is open.
Private Function FocusSearchAllPrompt(ByVal hMain As LongPtr, ByRef promptFound As Boolean) As Boolean
    Dim i As Long, v As Variant, h As LongPtr, yesButton As LongPtr
    Dim fg As LongPtr, fgTid As Long, dlgTid As Long, myTid As Long
    Dim dummy As Long, attachedFg As Boolean, attachedDlg As Boolean, attempt As Long

    For i = 1 To 20
        CollectResponsaWindows
        For Each v In mWins
            h = CLngPtr(v)
            If IsInfoOrResultModal(h) Then
                If FindButton(h, TxtYes()) <> 0 And FindButton(h, TxtNo()) <> 0 Then
                    promptFound = True
                    yesButton = FindButton(h, TxtYes())
                    Exit For
                End If
            End If
        Next
        If promptFound Then Exit For
        Sleep 100
        DoEvents
    Next
    If Not promptFound Then Exit Function

    fg = GetForegroundWindow()
    fgTid = GetWindowThreadProcessId(fg, dummy)
    dlgTid = GetWindowThreadProcessId(h, dummy)
    myTid = GetCurrentThreadId()

    If fgTid <> 0 And fgTid <> myTid Then
        attachedFg = (AttachThreadInput(myTid, fgTid, 1) <> 0)
    End If
    If dlgTid <> 0 And dlgTid <> myTid And dlgTid <> fgTid Then
        attachedDlg = (AttachThreadInput(myTid, dlgTid, 1) <> 0)
    End If

    On Error GoTo CleanUp
    For attempt = 1 To 3
        BringWindowToTop hMain
        BringWindowToTop h
        SetForegroundWindow h
        SetFocus yesButton
        If GetForegroundWindow() = h Then
            FocusSearchAllPrompt = True
            Exit For
        End If
        Sleep 100
    Next

CleanUp:
    If attachedDlg Then AttachThreadInput myTid, dlgTid, 0
    If attachedFg Then AttachThreadInput myTid, fgTid, 0
End Function

'==============================================================================
' Bring RESPONSA or its citation dialog to the foreground. Temporarily attach
' the input queues of the foreground window, Word, and the target dialog.
'==============================================================================
Private Function BringCitationToFront(ByVal hMain As LongPtr, ByVal hDlg As LongPtr, ByVal hEdit As LongPtr) As Boolean
    Dim fg As LongPtr, fgTid As Long, dlgTid As Long, myTid As Long
    Dim dummy As Long, attachedFg As Boolean, attachedDlg As Boolean, attempt As Long

    If IsIconic(hMain) <> 0 Then ShowWindow hMain, SW_RESTORE
    ShowWindow hDlg, SW_SHOW

    fg = GetForegroundWindow()
    fgTid = GetWindowThreadProcessId(fg, dummy)
    dlgTid = GetWindowThreadProcessId(hDlg, dummy)
    myTid = GetCurrentThreadId()

    If fgTid <> 0 And fgTid <> myTid Then
        attachedFg = (AttachThreadInput(myTid, fgTid, 1) <> 0)
    End If
    If dlgTid <> 0 And dlgTid <> myTid And dlgTid <> fgTid Then
        attachedDlg = (AttachThreadInput(myTid, dlgTid, 1) <> 0)
    End If

    On Error GoTo CleanUp
    For attempt = 1 To 3
        BringWindowToTop hMain
        BringWindowToTop hDlg
        SetForegroundWindow hDlg
        SetFocus hEdit
        If GetForegroundWindow() = hDlg Then
            BringCitationToFront = True
            Exit For
        End If
        Sleep 100
    Next

CleanUp:
    If attachedDlg Then AttachThreadInput myTid, dlgTid, 0
    If attachedFg Then AttachThreadInput myTid, fgTid, 0
End Function

Private Sub BringToFront(ByVal hMain As LongPtr)
    Dim hTop As LongPtr, fg As LongPtr
    Dim fgTid As Long, tgtTid As Long, myTid As Long, dummy As Long

    hTop = GetLastActivePopup(hMain)
    If hTop = 0 Then hTop = hMain
    If IsIconic(hMain) <> 0 Then ShowWindow hMain, SW_RESTORE

    fg = GetForegroundWindow()
    fgTid = GetWindowThreadProcessId(fg, dummy)
    tgtTid = GetWindowThreadProcessId(hMain, dummy)
    myTid = GetCurrentThreadId()

    If fgTid <> 0 And fgTid <> myTid Then AttachThreadInput myTid, fgTid, 1
    If tgtTid <> 0 And tgtTid <> myTid Then AttachThreadInput myTid, tgtTid, 1

    BringWindowToTop hMain
    If FOCUS_STRONG Then SwitchToThisWindow hTop, 1
    SetForegroundWindow hTop

    If tgtTid <> 0 And tgtTid <> myTid Then AttachThreadInput myTid, tgtTid, 0
    If fgTid <> 0 And fgTid <> myTid Then AttachThreadInput myTid, fgTid, 0
End Sub

'==============================================================================
' Hebrew captions from code points
'==============================================================================
Private Function Heb(ParamArray codes() As Variant) As String
    Dim i As Long, s As String
    For i = LBound(codes) To UBound(codes)
        s = s & ChrW(CLng(codes(i)))
    Next
    Heb = s
End Function

Private Function TxtSearch() As String: TxtSearch = Heb(1489, 1510, 1506, 32, 1495, 1497, 1508, 1493, 1513): End Function
Private Function TxtBrowseTitle() As String: TxtBrowseTitle = Heb(1506, 1497, 1493, 1503): End Function
Private Function TxtOk() As String: TxtOk = Heb(1488, 1497, 1513, 1493, 1512): End Function
Private Function TxtYes() As String: TxtYes = Heb(1499, 1503): End Function
Private Function TxtNo() As String: TxtNo = Heb(1500, 1488): End Function
Private Function TxtCancel() As String: TxtCancel = Heb(1489, 1497, 1496, 1493, 1500): End Function
Private Function TxtInfoTitle() As String: TxtInfoTitle = Heb(1502, 1497, 1491, 1506): End Function
Private Function TxtResultsWord() As String: TxtResultsWord = Heb(1514, 1493, 1510, 1488, 1493, 1514): End Function
Private Function TxtEasyTitle() As String: TxtEasyTitle = Heb(1495, 1497, 1508, 1493, 1513, 32, 1511, 1500): End Function
Private Function TxtAdvTitle() As String: TxtAdvTitle = Heb(1495, 1497, 1508, 1493, 1513, 32, 1502, 1514, 1511, 1491, 1501): End Function
