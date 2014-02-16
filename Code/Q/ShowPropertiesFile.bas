Attribute VB_Name = "ShowPropertiesFile"
Type SHELLEXECUTEINFO
   cbSize As Long
   fMask As Long
   hwnd As Long
   lpVerb As String
   lpFile As String
   lpParameters As String
   lpDirectory As String
   nShow As Long
   hInstApp As Long
   lpIDList As Long
   lpClass As String
   hkeyClass As Long
   dwHotKey As Long
   hIcon As Long
   hProcess As Long
End Type

Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Public Const BIF_RETURNONLYFSDIRS = 1
Public Const MAX_PATH = 260

Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long


Public Const SEE_MASK_INVOKEIDLIST = &HC
Public Const SEE_MASK_NOCLOSEPROCESS = &H40
Public Const SEE_MASK_FLAG_NO_UI = &H400
Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long
Public Sub ShowProps(FileName As String, OwnerhWnd As Long)
   Dim SEI As SHELLEXECUTEINFO
   Dim r As Long
   With SEI
      .cbSize = Len(SEI)
      .fMask = SEE_MASK_NOCLOSEPROCESS Or _
      SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
      .hwnd = OwnerhWnd
      .lpVerb = "properties"
      .lpFile = FileName
      .lpParameters = vbNullChar
      .lpDirectory = vbNullChar
      .nShow = 0
      .hInstApp = 0
      .lpIDList = 0
   End With
   r = ShellExecuteEX(SEI)
End Sub


Function BrowserFolder(Noidung As String) As String
Dim iNull As Integer, lpIDList As Long, lResult As Long
    Dim sPath As String, udtBI As BrowseInfo

    With udtBI
        .hWndOwner = frmMain.hwnd
        .lpszTitle = lstrcat(Noidung, "")
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        SHGetPathFromIDList lpIDList, sPath
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If
    BrowserFolder = sPath
End Function

