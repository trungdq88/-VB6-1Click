Attribute VB_Name = "mUniMsgBox"
Option Explicit
'
' Module:       Unicode Message Box (mUniMsgBox.bas)
' Yeu cau:      Ham VniStrToUni (mUniFunc.bas)
' Nguoi viet:   thuongall
' Email:        thuongall@yahoo.com
' Website:      www.caulacbovb.com
' Su dung:      Call UniMsgBox(VniStrToUni("Ca6u la5c bo65 VB"), vbInformation, VniStrToUni("Cha2o ba5n!"), Me.hWnd)
'
Private hDlgHook As Long

Private Const FONT_FACE = "Tahoma"

Private Const WH_CBT = 5
Private Const HCBT_ACTIVATE = 5
Private Const WM_SETFONT = &H30

Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal h As Long, ByVal w As Long, ByVal E As Long, ByVal O As Long, ByVal w As Long, ByVal i As Long, ByVal U As Long, ByVal s As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal ParenthWnd As Long, ByVal ChildhWnd As Long, ByVal ClassName As String, ByVal Caption As String) As Long
Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hMod As Long, ByVal dwThreadId As Long) As Long
Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Declare Function SetWindowTextW Lib "user32" (ByVal hWnd As Long, ByVal lpString As Long) As Long
Declare Function MessageBoxW Lib "user32.dll" (ByVal hWnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long

Function UniMsgBox(strText As String, Optional iButtons As VbMsgBoxStyle = vbOKOnly, Optional strTitle As String, Optional hWnd As Long = &H0) As VbMsgBoxResult
    hDlgHook = SetWindowsHookEx(WH_CBT, AddressOf HookProc, App.hInstance, GetCurrentThreadId())
    UniMsgBox = MessageBoxW(hWnd, StrPtr(strText), StrPtr(strTitle), iButtons)
End Function

Private Function HookProc(ByVal ncode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim hStatic1 As Long, hStatic2 As Long, hButton As Long, hFont As Long
    HookProc = CallNextHookEx(hDlgHook, ncode, wParam, lParam)
    If ncode = HCBT_ACTIVATE Then
        hFont = CreateFont(13, 0, 0, 0, 500, 0, 0, 0, 0, 0, 0, 0, 0, FONT_FACE)
   
        hStatic1 = FindWindowEx(wParam, 0&, "Static", vbNullString)
        hStatic2 = FindWindowEx(wParam, hStatic1, "Static", vbNullString)
        If hStatic2 = 0 Then hStatic2 = hStatic1
        SendMessage hStatic2, WM_SETFONT, hFont, ByVal 1&
   
        hButton = FindWindowEx(wParam, 0&, "Button", "OK")
        SendMessage hButton, WM_SETFONT, hFont, 0
        SetWindowTextW hButton, StrPtr(ChrW(&H110) & "óng")
   
        hButton = FindWindowEx(wParam, 0&, "Button", "&Yes")
        SendMessage hButton, WM_SETFONT, hFont, 0
        SetWindowTextW hButton, StrPtr("Có")
   
        hButton = FindWindowEx(wParam, 0&, "Button", "&No")
        SendMessage hButton, WM_SETFONT, hFont, 0
        SetWindowTextW hButton, StrPtr("Không")
   
         hButton = FindWindowEx(wParam, 0&, "Button", "&Retry")
        SendMessage hButton, WM_SETFONT, hFont, 0
        SetWindowTextW hButton, StrPtr("Th" & ChrW(&H1EED) & " l" & ChrW(&H1EA1) & "i")
   
        hButton = FindWindowEx(wParam, 0&, "Button", "Cancel")
        SendMessage hButton, WM_SETFONT, hFont, 0
        SetWindowTextW hButton, StrPtr("Thoát")
       
       UnhookWindowsHookEx hDlgHook
    End If
End Function

