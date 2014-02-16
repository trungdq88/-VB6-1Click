Attribute VB_Name = "DiaLog1"
Option Explicit
Const FW_NORMAL = 400
Const DEFAULT_CHARSET = 1
Const OUT_DEFAULT_PRECIS = 0
Const CLIP_DEFAULT_PRECIS = 0
Const DEFAULT_QUALITY = 0
Const DEFAULT_PITCH = 0
Const FF_ROMAN = 16
Const CF_PRINTERFONTS = &H2
Const CF_SCREENFONTS = &H1
Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Const CF_EFFECTS = &H100&
Const CF_FORCEFONTEXIST = &H10000
Const CF_INITTOLOGFONTSTRUCT = &H40&
Const CF_LIMITSIZE = &H2000&
Const REGULAR_FONTTYPE = &H400
Const LF_FACESIZE = 32
Const CCHDEVICENAME = 32
Const CCHFORMNAME = 32
Const GMEM_MOVEABLE = &H2
Const GMEM_ZEROINIT = &H40
Const DM_DUPLEX = &H1000&
Const DM_ORIENTATION = &H1&
Const PD_PRINTSETUP = &H40
Const PD_DISABLEPRINTTOFILE = &H80000
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Type CHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName As String * 31
End Type
Private Type CHOOSEFONT
        lStructSize As Long
        hwndOwner As Long
        hDC As Long
        lpLogFont As Long
        iPointSize As Long
        flags As Long
        rgbColors As Long
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
        hInstance As Long
        lpszStyle As String
        nFontType As Integer
        MISSING_ALIGNMENT As Integer
        nSizeMin As Long
        nSizeMax As Long
End Type
Private Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function CHOOSEFONT Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONT) As Long
Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GlobalLock Lib "Kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "Kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "Kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "Kernel32" (ByVal hMem As Long) As Long
Dim OFName As OPENFILENAME
Dim mHwnd As Long
Dim CustomColors() As Byte
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpOperation _
    As String, ByVal lpFile As String, ByVal lpParameters _
    As String, ByVal lpDirectory As String, ByVal nShowCmd _
    As Long) As Long
Private Const BIF_STATUSTEXT = &H4&
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260
Private Const WM_USER = &H400
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELCHANGED = 2
Private Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Private Const BFFM_SETSELECTION = (WM_USER + 102)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "Kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Type BrowseInfo
  hwndOwner      As Long
  pIDLRoot       As Long
  pszDisplayName As Long
  lpszTitle      As Long
  ulFlags        As Long
  lpfnCallback   As Long
  lParam         As Long
  iImage         As Long
End Type
Private m_CurrentDirectory As String
Public mFontName As String
Public mFontsize As Integer
Public mBold As Boolean
Public mItalic As Boolean
Public mUnderline As Boolean
Public mStrikethru As Boolean
Public mFontColor As Long
Public mFilterIndex As Integer


Public Function BrowseForFolder(Optional StartDir As String, Optional mTitle As String) As String
  Dim lpIDList As Long
  Dim szTitle As String
  Dim sBuffer As String
  Dim tBrowseInfo As BrowseInfo
  If StartDir = "" Then StartDir = "c:\"
  If mTitle = "" Then mTitle = App.Title + " : Select a Folder"
  m_CurrentDirectory = StartDir & vbNullChar
  szTitle = mTitle
  With tBrowseInfo
    .hwndOwner = mHwnd
    .lpszTitle = lstrcat(szTitle, "")
    .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN + BIF_STATUSTEXT
    .lpfnCallback = GetAddressofFunction(AddressOf BrowseCallbackProc)
  End With
  lpIDList = SHBrowseForFolder(tBrowseInfo)
  If (lpIDList) Then
    sBuffer = Space(MAX_PATH)
    SHGetPathFromIDList lpIDList, sBuffer
    sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    BrowseForFolder = sBuffer
  Else
    BrowseForFolder = ""
  End If
End Function
Private Function BrowseCallbackProc(ByVal Hwnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
Dim lpIDList As Long
Dim ret As Long
Dim sBuffer As String
On Error Resume Next
Select Case uMsg
  Case BFFM_INITIALIZED
    Call SendMessage(Hwnd, BFFM_SETSELECTION, 1, m_CurrentDirectory)
  Case BFFM_SELCHANGED
    sBuffer = Space(MAX_PATH)
    ret = SHGetPathFromIDList(lp, sBuffer)
    If ret = 1 Then
      Call SendMessage(Hwnd, BFFM_SETSTATUSTEXT, 0, sBuffer)
    End If
End Select
BrowseCallbackProc = 0
End Function
Private Function GetAddressofFunction(add As Long) As Long
  GetAddressofFunction = add
End Function


Public Sub MyFind(initDir As String)
ShellExecute 0, "find", initDir, vbNullString, vbNullString, 5

End Sub
Public Sub InitCmnDlg(myhwnd As Long)
'Required to use custom colors
ReDim CustomColors(0 To 16 * 4 - 1) As Byte
Dim i As Integer
For i = LBound(CustomColors) To UBound(CustomColors)
    CustomColors(i) = 0
Next i
'need a window handle to run the functions
mHwnd = myhwnd
End Sub

Public Function ShowColor() As Long
    Dim cc As CHOOSECOLOR
    Dim Custcolor(16) As Long
    Dim lReturn As Long
    cc.lStructSize = Len(cc)
    cc.hwndOwner = mHwnd
    cc.hInstance = App.hInstance
    cc.lpCustColors = StrConv(CustomColors, vbUnicode)
    cc.flags = 0
    If CHOOSECOLOR(cc) <> 0 Then
        ShowColor = cc.rgbResult
        CustomColors = StrConv(cc.lpCustColors, vbFromUnicode)
    Else
        ShowColor = -1
    End If
End Function
Public Function ShowOpen(Optional mFilter As String, Optional mflags As Long, Optional mInitDir As String, Optional mTitle As String) As String
    If mInitDir = "" Then mInitDir = "c:\"
    If mFilter = "" Then mFilter = "All Files (*.*)" + Chr(0) + "*.*" + Chr(0)
    If mTitle = "" Then mTitle = App.Title
    OFName.lStructSize = Len(OFName)
    OFName.hwndOwner = mHwnd
    OFName.hInstance = App.hInstance
    OFName.lpstrFilter = mFilter
    OFName.lpstrFile = Space$(254)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = Space$(254)
    OFName.nMaxFileTitle = 255
    OFName.lpstrInitialDir = mInitDir
    OFName.lpstrTitle = mTitle
    OFName.flags = mflags
    If GetOpenFileName(OFName) Then
        ShowOpen = StripTerminator(OFName.lpstrFile)
    Else
        ShowOpen = ""
    End If
End Function
Public Function ShowFont() As Boolean
    Dim cf As CHOOSEFONT, lfont As LOGFONT, hMem As Long, pMem As Long
    Dim RetVal As Long
    mFontName = ""
    lfont.lfHeight = 0
    lfont.lfWidth = 0
    lfont.lfEscapement = 0
    lfont.lfOrientation = 0
    lfont.lfWeight = FW_NORMAL
    lfont.lfCharSet = DEFAULT_CHARSET
    lfont.lfOutPrecision = OUT_DEFAULT_PRECIS
    lfont.lfClipPrecision = CLIP_DEFAULT_PRECIS
    lfont.lfQuality = DEFAULT_QUALITY
    lfont.lfPitchAndFamily = DEFAULT_PITCH Or FF_ROMAN
    lfont.lfFaceName = "Times New Roman" & vbNullChar
    hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(lfont))
    pMem = GlobalLock(hMem)
    CopyMemory ByVal pMem, lfont, Len(lfont)
    cf.lStructSize = Len(cf)
    cf.hwndOwner = frmMain.Hwnd
    cf.hDC = Printer.hDC
    cf.lpLogFont = pMem
    cf.iPointSize = 120
    cf.flags = CF_BOTH Or CF_EFFECTS Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT Or CF_LIMITSIZE
    cf.rgbColors = RGB(0, 0, 0)
    cf.nFontType = REGULAR_FONTTYPE
    cf.nSizeMin = 10
    cf.nSizeMax = 72
    RetVal = CHOOSEFONT(cf)
    If RetVal <> 0 Then
        ShowFont = True
        CopyMemory lfont, ByVal pMem, Len(lfont)
        mFontName = Left(lfont.lfFaceName, InStr(lfont.lfFaceName, vbNullChar) - 1)
        mBold = False
        mItalic = False
        mUnderline = False
        mStrikethru = False
        mFontsize = cf.iPointSize / 10
        If lfont.lfItalic = 255 Then mItalic = True
        If lfont.lfUnderline = 255 Then mUnderline = True
        If lfont.lfWeight = 700 Then mBold = True
        If lfont.lfStrikeOut = 255 Then mStrikethru = True
        mFontColor = cf.rgbColors
    Else
        ShowFont = False
    End If
    RetVal = GlobalUnlock(hMem)
    RetVal = GlobalFree(hMem)
End Function
Public Function ShowSave(Optional mFilter As String, Optional mflags As Long, Optional mInitDir As String, Optional mTitle As String) As String
    If mInitDir = "" Then mInitDir = "c:\"
    If mFilter = "" Then mFilter = "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    If mTitle = "" Then mTitle = App.Title
    OFName.lStructSize = Len(OFName)
    OFName.hwndOwner = mHwnd
    OFName.hInstance = App.hInstance
    OFName.lpstrFilter = mFilter
    OFName.lpstrFile = Space$(254)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = Space$(254)
    OFName.nMaxFileTitle = 255
    OFName.lpstrInitialDir = mInitDir
    OFName.lpstrTitle = mTitle
    OFName.flags = mflags
    
    If GetSaveFileName(OFName) Then
        mFilterIndex = OFName.nFilterIndex
        ShowSave = StripTerminator(OFName.lpstrFile)
    Else
        ShowSave = ""
    End If
End Function
Public Function StripTerminator(ByVal strString As String) As String
'gets rid of anything not required returned by API calls
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function







