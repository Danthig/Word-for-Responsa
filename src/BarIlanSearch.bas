Attribute VB_Name = "BarIlanSearch"
Option Explicit
'==============================================================================
' Search the text selected in Word inside Bar-Ilan Responsa (RESPONSA.exe)
'   SearchSelectionInBarIlan          - plain words, in the "easy search" window
'   SearchSelectionAdvancedInBarIlan  - "advanced search" window: every word wrapped with
'                                       prefix/suffix, with a distance in front, e.g.
'                                       30: #!word1# #!word2# #!word3#
' Non-contiguous (Ctrl) selections are supported: all selected areas are searched.
' Window IDs / commands taken from helper/lib/src/native/responsa_search_automation.dart
' Word 2010+ (VBA7). Hebrew strings are built with ChrW so the file stays ASCII.
'==============================================================================

'---- settings ----------------------------------------------------------------
Private Const RESPONSA_PATH As String = ""      ' optional: full path to RESPONSA.exe
Private Const CMD_OPEN_SEARCH As Long = 32857
Private Const ID_SEARCH_EDIT As Long = 1233
Private Const ID_ADVANCED_BTN As Long = 1209    ' button that switches to "advanced search"
Private Const ID_ADVANCED_ONLY As Long = 1065   ' control that exists only in advanced mode
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
Private Const BM_CLICK As Long = &HF5
Private Const SMTO_ABORTIFHUNG As Long = 2
Private Const GW_OWNER As Long = 4
Private Const SW_RESTORE As Long = 9
Private Const PROCESS_QUERY_LIMITED_INFORMATION As Long = &H1000
Private Const CF_UNICODETEXT As Long = 13
Private Const GMEM_MOVEABLE As Long = 2

Private Declare PtrSafe Function EnumWindows Lib "user32" (ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function EnumChildWindows Lib "user32" (ByVal hWndParent As LongPtr, ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As LongPtr, ByRef lpdwProcessId As Long) As Long
Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function IsIconic Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
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

Private Sub RunSearch(ByVal advanced As Boolean)
    Dim q As String
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
        MsgBox "RESPONSA is not running (and RESPONSA_PATH is not set).", vbExclamation
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
    Sleep 300                                   ' let the dialog take the text
    PostMessageW hBtn, BM_CLICK, 0, 0           ' async: the button opens a modal
    Sleep 200
    BringToFront hMain                          ' exactly once, no verification, no retry
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
    Dim h As LongPtr, t0 As Double
    h = FindMainWindow()
    If h <> 0 Then
        EnsureResponsaRunning = h
        Exit Function
    End If
    If Len(RESPONSA_PATH) = 0 Then Exit Function
    Shell """" & RESPONSA_PATH & """", vbNormalFocus
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
    Dim r As LongPtr
    SetText = (SendMessageTimeoutW(h, WM_SETTEXT, 0, StrPtr(s), SMTO_ABORTIFHUNG, 2000, r) <> 0)
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
' free text) and shows only the selected one; a hidden one runs a search fine.
' So the dialog is picked by its title - mode 0 = easy, 1 = advanced, 2 = any.
'==============================================================================
Private Function FindSearchDialog(ByVal mode As Long, ByRef btn As LongPtr, ByRef edt As LongPtr) As LongPtr
    Dim pass As Long, v As Variant, h As LongPtr, b As LongPtr, e As LongPtr
    Dim title As String, isAdv As Boolean, ok As Boolean
    CollectResponsaWindows
    For pass = 1 To 2                        ' visible first, hidden as fallback
        For Each v In mWins
            h = CLngPtr(v)
            If pass = 2 Or IsWindowVisible(h) <> 0 Then
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
    Next
End Function

' open the search dialogs (command resent every few seconds) and return the one
' for the wanted mode
Private Function EnsureSearchDialog(ByVal hMain As LongPtr, ByVal advanced As Boolean, ByRef btn As LongPtr, ByRef edt As LongPtr) As LongPtr
    Dim h As LongPtr, t0 As Double, tSent As Double, mode As Long
    mode = IIf(advanced, 1, 0)

    h = FindSearchDialog(mode, btn, edt)
    If h <> 0 Then EnsureSearchDialog = h: Exit Function

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
        ' dialogs exist but none matched by title (other language edition?)
        If Timer - t0 >= FALLBACK_AFTER_SEC Then
            h = FindSearchDialog(2, btn, edt)
            If h <> 0 Then Exit Do
        End If
    Loop While Timer - t0 < DIALOG_WAIT_SEC
    If h = 0 Then Exit Function

    If advanced Then
        EnsureSearchDialog = EnsureAdvancedMode(h, btn, edt)
    Else
        EnsureSearchDialog = h
    End If
End Function

' fallback: switch the visible dialog to "advanced search" with its mode button
Private Function EnsureAdvancedMode(ByVal hDlg As LongPtr, ByRef btn As LongPtr, ByRef edt As LongPtr) As LongPtr
    Dim b As LongPtr, d As LongPtr, t0 As Double
    If ChildById(hDlg, ID_ADVANCED_ONLY) <> 0 Then EnsureAdvancedMode = hDlg: Exit Function

    b = ChildById(hDlg, ID_ADVANCED_BTN)
    If b = 0 Then Exit Function
    PostMessageW b, BM_CLICK, 0, 0

    t0 = Timer
    Do
        Sleep 150
        DoEvents
        d = FindSearchDialog(1, btn, edt)
        If d <> 0 Then EnsureAdvancedMode = d: Exit Function
    Loop While Timer - t0 < 6
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

'==============================================================================
' Bring RESPONSA (and the modal on top of it) to the foreground - one attempt.
' Attaches to the input queues of the current foreground thread and of RESPONSA,
' which is what lets SetForegroundWindow succeed.
'==============================================================================
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
Private Function TxtOk() As String: TxtOk = Heb(1488, 1497, 1513, 1493, 1512): End Function
Private Function TxtCancel() As String: TxtCancel = Heb(1489, 1497, 1496, 1493, 1500): End Function
Private Function TxtInfoTitle() As String: TxtInfoTitle = Heb(1502, 1497, 1491, 1506): End Function
Private Function TxtResultsWord() As String: TxtResultsWord = Heb(1514, 1493, 1510, 1488, 1493, 1514): End Function
Private Function TxtEasyTitle() As String: TxtEasyTitle = Heb(1495, 1497, 1508, 1493, 1513, 32, 1511, 1500): End Function
Private Function TxtAdvTitle() As String: TxtAdvTitle = Heb(1495, 1497, 1508, 1493, 1513, 32, 1502, 1514, 1511, 1491, 1501): End Function
