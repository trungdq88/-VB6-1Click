Attribute VB_Name = "SysMenu"
Option Explicit

Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const WM_SYSCOMMAND = &H112
Public Const MF_SEPARATOR = &H800&
Public Const MF_STRING = &H0&
Public Const GWL_WNDPROC = (-4)
Public Const IDM_ABOUT As Long = 1010
Public lProcOld As Long

Public Function SysMenuHandler(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

If iMsg = WM_SYSCOMMAND Then
    If wParam = IDM_ABOUT Then
        'MsgBox "www.caulacbovb.com", vbInformation, "About"
        frmAbout.Show
        Exit Function
    End If
End If
SysMenuHandler = CallWindowProc(lProcOld, hWnd, iMsg, wParam, lParam)

End Function
Public Function SubClass(FormName As Form)

Dim lhSysMenu As Long, lRet As Long

    lhSysMenu = GetSystemMenu(FormName.hWnd, 0&)
    lRet = AppendMenu(lhSysMenu, MF_SEPARATOR, 0&, _
    vbNullString)
    lRet = AppendMenu(lhSysMenu, MF_STRING, _
    IDM_ABOUT, "About...")
    FormName.Show
    lProcOld = SetWindowLong(FormName.hWnd, GWL_WNDPROC, AddressOf SysMenuHandler)

End Function

