Attribute VB_Name = "modUniInputbox"


Private Const GWL_WNDPROC = (-4&)
Private Const WH_CBT As Long = &H5
Private Const HCBT_ACTIVATE As Long = &H5
Public Const WM_SETTEXT = &HC
Public Const WM_SETFONT = &H30
Public Const NV_INPUTBOX As Long = &H5000&
Private Const EM_SETPASSWORDCHAR = &HCC
 
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal w As Long, ByVal E As Long, ByVal O As Long, ByVal w As Long, ByVal i As Long, ByVal U As Long, ByVal s As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal f As String) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd&, ByVal nIndex&, ByVal dwNewLong&) As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal CodeNo As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal ParenthWnd As Long, ByVal ChildhWnd As Long, ByVal ClassName As String, ByVal Caption As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function MessageBoxW Lib "user32.dll" (ByVal hwnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long
Private Declare Function SetWindowTextW Lib "user32" (ByVal hwnd As Long, ByVal lpString As Long) As Long
Private Declare Function DefWindowProcW Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetWindowTextW Lib "user32.dll" (ByVal hwnd As Long, ByVal lpString As Long, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLengthW Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetTimer& Lib "user32" (ByVal hwnd&, ByVal nIDEvent&, ByVal uElapse&, ByVal lpTimerFunc&)
Private Declare Function KillTimer& Lib "user32" (ByVal hwnd&, ByVal nIDEvent&)
Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
 
Private pHook2 As Long, pHook3 As Long, hEdit As Long, hIdEvent As Long, UsePass As Boolean
Private sStatic As String, sDefault As String, sTitle As String, sInput As String, txt As String
 
Private Function InputHookProc(ByVal ncode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim hStatic1 As Long, hStatic2 As Long, hButton As Long, hFont As Long
InputHookProc = CallNextHookEx(pHook2, ncode, wParam, lParam)
If ncode = HCBT_ACTIVATE Then
   hFont = CreateFont(13, 0, 0, 0, 500, 0, 0, 0, 0, 0, 0, 0, 0, "Tahoma")
   
   hStatic1 = FindWindowEx(wParam, 0&, "Static", vbNullString)
   hStatic2 = FindWindowEx(wParam, hStatic1, "Static", vbNullString)
   If hStatic2 = 0 Then hStatic2 = hStatic1
   SendMessage hStatic2, WM_SETFONT, hFont, ByVal 1&
   DefWindowProcW hStatic2, WM_SETTEXT, &H0&, StrPtr(sStatic)
   DefWindowProcW wParam, WM_SETTEXT, &H0&, StrPtr(sTitle)
   
   hButton = FindWindowEx(wParam, 0&, "Button", "OK")
   SendMessage hButton, WM_SETFONT, hFont, ByVal 1&
   DefWindowProcW hButton, WM_SETTEXT, &H0&, StrPtr("Xác nh" & ChrW(7853) & "n")
   
   hButton = FindWindowEx(wParam, 0&, "Button", "Cancel")
   SendMessage hButton, WM_SETFONT, hFont, ByVal 1&
   DefWindowProcW hButton, WM_SETTEXT, &H0&, StrPtr("H" & ChrW(7911) & "y b" & ChrW(7887))
 
    hEdit = FindWindowEx(wParam, 0&, "Edit", "")
    SendMessage hEdit, WM_SETFONT, hFont, ByVal 1&
   
    If sDefault <> "" Then
    SetWindowTextW hEdit, StrPtr(sDefault) 'Khong ho tro Tieng Viet o Input Textbox khi Style = Windows Classic
    SendKeys "+{END}" 'Select text
    End If
     
    If UsePass Then SendMessage hEdit, EM_SETPASSWORDCHAR, Asc("*"), 0
   
    UnhookWindowsHookEx pHook3
End If
End Function
 
Public Function UniInputbox(ByVal Prompt As String, Optional ByVal Title As String = "", Optional ByVal Default As String = "", Optional ByVal Password As Boolean = False) As String
    pHook3 = SetWindowsHookEx(WH_CBT, AddressOf InputHookProc, App.hInstance, GetCurrentThreadId())
    UsePass = Password
    sStatic = VnToUni(Prompt)
    sDefault = VnToUni(Default)
    sTitle = VnToUni(Title)
    SetTimer 0, NV_INPUTBOX, 50, AddressOf TimerProc 'Lay du lieu Tieng Viet o Input Text Box
    txt = Inputbox(sStatic, sTitle, sDefault)
    KillTimer 0, hIdEvent
    If txt <> "" Then UniInputbox = StripNulls(sInput)
End Function
 
Public Sub TimerProc(ByVal hwnd&, ByVal uMsg&, ByVal idEvent&, ByVal dwTime&)
If hEdit <> 0 Then sInput = GetUniText(hEdit) 'Copy lien tuc ^^!
hIdEvent = idEvent
End Sub
 
Private Function GetUniText(ByVal hwnd As Long) As String
Dim lLen As Long, sBuf As String
lLen = 1 + GetWindowTextLengthW(hwnd)
If (lLen > 1) Then
    sBuf = String$(lLen, 0)
    GetWindowTextW hwnd, StrPtr(sBuf), lLen
    GetUniText = (sBuf)
Else
    GetUniText = vbNullString
End If
End Function
 
Private Function StripNulls(ByVal sString As String) As String
Dim lPos As Long
    lPos = InStr(sString, vbNullChar)
    If (lPos = 1) Then
        StripNulls = vbNullString
    ElseIf (lPos > 1) Then
        StripNulls = Left$(sString, lPos - 1)
        Exit Function
    End If
    StripNulls = sString
End Function
 
'Code convert TCVN3 -> Unicode by TruongPhu
Public Function VnToUni(str As String) As String
Dim i&, arrUNI() As String, sUni$, ABC$, UNI$
ABC = "¸µ¶·¹¨¾»¼½Æ©ÊÇÈÉËÐÌÎÏÑªÕÒÓÔÖÝ×ØÜÞãßáâä«èåæçé¬íêëìîóïñòô­øõö÷ùýúûüþ®¸µ¶·¹¡¾»¼½Æ¢ÊÇÈÉËÐÌÎÏÑ£ÕÒÓÔÖÝ×ØÜÞãßáâä¤èåæçé¥íêëìîóïñòô¦øõö÷ùýúûüþ§"
UNI = "225,224,7843,227,7841,259,7855,7857,7859,7861,7863,226,7845,7847,7849,7851,7853,233,232,7867,7869,7865,234,7871,7873,7875,7877,7879,237,236,7881,297,7883,243,242,7887,245,7885,244,7889,7891,7893,7895,7897,417,7899,7901,7903,7905,7907,250,249,7911,361,7909,432,7913,7915,7917,7919,7921,253,7923,7927,7929,7925,273,225,224,7843,227,7841,258,7855,7857,7859,7861,7863,194,7845,7847,7849,7851,7853,233,232,7867,7869,7865,202,7871,7873,7875,7877,7879,237,236,7881,297,7883,243,242,7887,245,7885,212,7889,7891,7893,7895,7897,416,7899,7901,7903,7905,7907,250,249,7911,361,7909,431,7913,7915,7917,7919,7921,253,7923,7927,7929,7925,272"
arrUNI = Split(UNI, ",")
For i = 1 To Len(str$)
If InStr(ABC, Mid(str$, i, 1)) > 0 Then
 sUni = sUni & ChrW(arrUNI(InStr(ABC, Mid(str$, i, 1)) - 1))
 Else
 sUni = sUni & Mid(str$, i, 1)
 End If
Next
VnToUni = sUni
End Function
