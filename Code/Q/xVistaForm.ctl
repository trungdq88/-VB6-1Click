VERSION 5.00
Begin VB.UserControl xVistaForm 
   Alignable       =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2310
   MaskColor       =   &H00FF00FF&
   ScaleHeight     =   570
   ScaleWidth      =   2310
End
Attribute VB_Name = "xVistaForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    '****************************************************************
    ' Project:      Creates an Ownerdrawn Vista Style Form control
    ' Programmer:   Alexander Mungall
    ' UserControl:  xVistaForm
    ' Email:        goober_mpc@hotmail.com
    '----------------------------------------------------------------
    ' xVistaForm Copyright© Alexander Mungall, All Rights Reserved
    ' Feel free to use this code for personal use in anyway you see
    ' fit, but please give credit where credit is due...
    ' It's all I ask.
    '****************************************************************
    Option Explicit
    
    ' Booleans
    Private bCloseButton As Boolean
    Private bCloseButtonClicked As Boolean
    Private bCloseButtonHover As Boolean
    Private bDisplayIcon As Boolean
    Private bFontBold As Boolean
    Private bFontItalic As Boolean
    Private bFontStrikeThru As Boolean
    Private bFontUnderline As Boolean
    Private bMaximiseButton As Boolean
    Private bMaximiseButtonClicked As Boolean
    Private bMaximiseButtonHover As Boolean
    Private bMinimiseButton As Boolean
    Private bMinimiseButtonClicked As Boolean
    Private bMinimiseButtonHover As Boolean
    Private bMouseClicked As Boolean
    Private bMouseOnForm As Boolean
    Private bPaintForm As Boolean
    Private bRightClick As Boolean
    Private bSystemTray As Boolean
    Private bTransparency As Boolean
    Private bUnloadForm As Boolean
    
    ' Controls
    Private imgFormPic As Image
    Private myForm As Form
    Private WithEvents moveForm As Form
Attribute moveForm.VB_VarHelpID = -1
    Private WithEvents picForm As PictureBox
Attribute picForm.VB_VarHelpID = -1
    Private WithEvents tmrMouseMove As Timer
Attribute tmrMouseMove.VB_VarHelpID = -1
    
    ' Doubles
    Private dFontSize As Double
    
    ' Enums
    Public Enum xVistaStyles
        VistaLite = 0
        VistaDark = 1
    End Enum
    Private xVisualStyles As xVistaStyles

    ' Integers
    Private i As Integer
    Private iHorizontal As Integer
    Private iNumControls As Integer
    Private iTransparency As Integer
    Private iVertical As Integer
    
    ' Longs
    Private Col As Long
    Private lBottomR As Long
    Private lBottomG As Long
    Private lBottomB As Long
    Private lButtonGradientBottom(7) As Long
    Private lButtonGradientBottomClicked(7) As Long
    Private lButtonGradientBottomHover(7) As Long
    Private lButtonGradientTop As Long
    Private lButtonGradientTopClicked As Long
    Private lButtonGradientTopHover As Long
    Private lButtonInnerBorder As Long
    Private lButtonOuterBorder As Long
    Private lCloseButtonGradientBottom(7) As Long
    Private lCloseButtonGradientBottomClicked(7) As Long
    Private lCloseButtonGradientBottomHover(7) As Long
    Private lCloseButtonGradientTop As Long
    Private lCloseButtonGradientTopClicked As Long
    Private lCloseButtonGradientTopHover As Long
    Private lCloseButtonInnerBorder As Long
    Private lCloseButtonOuterBorder As Long
    Private lTopR As Long
    Private lTopG As Long
    Private lTopB As Long
    Private lFormCaptionColor As Long
    Private lFormGradientBottom As Long
    Private lFormGradientTop As Long
    Private lFormInnerBorder As Long
    Private lFormMaxHeight As Long
    Private lFormMinHeight As Long
    Private lFormMaxWidth As Long
    Private lFormMinWidth As Long
    Private lFormMiddleBorder As Long
    Private lFormOuterBorder As Long
    Private lngReturnValue As Long
    Private lSysTrayMenu As Long
    
    ' Strings
    Private sFormCaption As String

    ' Move a Titleless Window
    Private Declare Function ReleaseCapture Lib "user32" () As Long
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Const HTCAPTION = 2
    Private Const WM_NCLBUTTONDOWN = &HA1
    Private Const WM_SYSCOMMAND = &H112
    Private Const SC_MOVE = &HF010
    Private Const WM_POPUPSYSTEMMENU = &H313
    
    ' Make a Semi Transparent Form
    Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
    Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hWnd As Long, ByVal hDCDst As Long, pptDst As Any, psize As Any, ByVal hDCSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Const G = (-20)
    Private Const LWA_COLORKEY = &H1
    Private Const LWA_ALPHA = &H2
    Private Const ULW_COLORKEY = &H1
    Private Const ULW_ALPHA = &H2
    Private Const ULW_OPAQUE = &H4
    Private Const WS_EX_LAYERED = &H80000
    
    ' Show a Form in the Taskbar
    Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
    Private Const GWL_EXSTYLE = (-20)
    Private Const WS_EX_APPWINDOW = &H40000
    Private Const SW_HIDE = 0
    Private Const SW_NORMAL = 1

    ' Types
    Private Type POINTAPI
        x As Long
        y As Long
    End Type
    
    ' Functions
    Private Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
    Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

    ' Local constants and variables declarations
    Const BorderPixels = 5
    Private TwipX As Single, TwipY As Single, BorderWidth As Single, BorderHeight As Single

    ' Type passed to Shell_NotifyIcon
    Private Type NotifyIconData
      Size As Long
      Handle As Long
      ID As Long
      Flags As Long
      CallBackMessage As Long
      Icon As Long
      Tip As String * 64
    End Type

    ' Region combine consts for making the Transparent areas
    Private Const RGN_AND = 1   'Combines an intersection
    Private Const RGN_OR = 2    'Creates a union of two regions
    Private Const RGN_XOR = 3   'Creations a union of two objects with the exception of overlapping
    Private Const RGN_DIFF = 4  'Combines two regions
    Private Const RGN_COPY = 5  'Copy a region

    ' Declarations for making the Transparent areas
    Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
    Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
    Private Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long
    Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
    
    Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
    Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
    Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
    
    'Our declarations for retrieving colors
    Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
    Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

    ' Constants for managing System Tray tasks
    Private Const AddIcon = &H0
    Private Const ModifyIcon = &H1
    Private Const DeleteIcon = &H2
    
    Private Const WM_MOUSEMOVE = &H200
    Private Const WM_LBUTTONDBLCLK = &H203
    Private Const WM_LBUTTONDOWN = &H201
    Private Const WM_LBUTTONUP = &H202
    
    Private Const WM_RBUTTONDBLCLK = &H206
    Private Const WM_RBUTTONDOWN = &H204
    Private Const WM_RBUTTONUP = &H205
    
    Private Const MessageFlag = &H1
    Private Const IconFlag = &H2
    Private Const TipFlag = &H4
    
    Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal Message As Long, Data As NotifyIconData) As Boolean
    
    Private Data As NotifyIconData

    ' Menu Functions
    Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu&, ByVal wFlags&, ByVal wIDNewItem&, ByVal lpNewItem$) As Long
    Private Declare Function CreatePopupMenu Lib "user32" () As Long
    Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu&) As Long
    'Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu&, ByVal wFlags&, ByVal X&, ByVal Y&, ByVal nReserved&, ByVal Hwnd&, ByVal lpRect&) As Long
    Private Declare Function EnableMenuItem Lib "user32" (ByVal hMenu&, ByVal wIDEnableItem&, ByVal wEnable&) As Long
    Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal B As Boolean, lpMenuItemInfo As MENUITEMINFO) As Boolean
    Private Declare Function SetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPos As Long) As Long
    Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
    Private Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal hWnd As Long, ByVal lptpm As Any) As Long
    
    ' Menu Constants, Enums & Events
    Private Const MFS_BYCOMMAND As Long = &H0&
    Private Const MFS_CHECKED As Long = &H8&
    Private Const MFS_DEFAULT As Long = &H1000&
    Private Const MFS_DISABLED As Long = &H2&
    Private Const MFS_ENABLED As Long = &H0
    Private Const MFS_GRAYED As Long = &H1&
    Private Const MFS_STRING As Long = &H0&
    Private Const MFS_SEPARATOR As Long = &H800&
    Private Const MIIM_CHECKMARKS As Long = &H8
    Private Const MIIM_DATA = &H20
    Private Const MIIM_ID = &H2
    Private Const MIIM_STATE As Long = &H1&
    Private Const MIIM_TYPE As Long = &H10&
    Private Const TPM_RETURNCMD As Long = &H100&
    Public Enum menuEStates
        xDisabled = 1
        xGrayed = 2
    End Enum

    Private Type MENUITEMINFO
        cbSize As Long
        fMask As Long
        fType As Long
        fState As Long
        wID As Long
        hSubMenu As Long
        hbmpChecked As Long
        hbmpUnchecked As Long
        dwItemData As Long
        dwTypeData As String
        cch As Long
    End Type

    Event Execute(ByVal ID As Long)
       
    '****************************************************************
    ' Gradient Code: Written by Mark Gordon (msg555)
    '----------------------------------------------------------------
    ' Copyright© Mark Gordon, All Rights Reserved
    '----------------------------------------------------------------
    Private Declare Function CreateBitmap Lib "gdi32.dll" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, ByRef lpBits As Any) As Long
    Private Declare Function GetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
    Private Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
    
    Private Const DIB_RGB_COLORS = 0&
    Private Const BI_RGB = 0&
    
    Private Type BITMAPINFOHEADER '40 bytes
       biSize As Long
       biWidth As Long
       biHeight As Long
       biPlanes As Integer
       biBitCount As Integer
       biCompression As Long
       biSizeImage As Long
       biXPelsPerMeter As Long
       biYPelsPerMeter As Long
       biClrUsed As Long
       biClrImportant As Long
    End Type
    
    Private Type RGBQUAD
       rgbBlue As Byte
       rgbGreen As Byte
       rgbRed As Byte
       rgbReserved As Byte
    End Type
    
    Private Type BITMAPINFO
      bmiHeader As BITMAPINFOHEADER
      bmiColors As RGBQUAD
    End Type
    
    Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PictDesc, riid As Guid, ByVal fPictureOwnsHandle As Long, iPic As StdPicture) As Long
    Private Type PictDesc
        cbSizeofStruct As Long
        picType As Long
        hImage As Long
        xExt As Long
        yExt As Long
    End Type
    Private Type Guid
        Data1 As Long
        Data2 As Integer
        Data3 As Integer
        Data4(0 To 7) As Byte
    End Type
    
    Private Enum Blends
        RGBBlend = 0
        HSLBlend = 1
    End Enum

' ham in text unicode
Private Declare Function TextOutWP Lib "gdi32" Alias "TextOutW" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As Long, ByVal nCount As Long) As Long

Private Function CreateGradient(Width As Long, Height As Long, LeftToRight As Boolean, LeftTopColor As Long, RightBottomColor As Long, BlendType As Blends) As StdPicture
    Dim hBmp As Long, Bits() As Byte
    Dim RS As Byte, GS As Byte, BS As Byte 'Start RGB
    Dim RE As Byte, GE As Byte, BE As Byte 'End RGB
    Dim HS As Single, SS As Single, LS As Single 'Start HSL
    Dim HE As Single, SE As Single, LE As Single 'End HSL
    Dim rc As Byte, GC As Byte, BC As Byte 'Current iteration RGB
    Dim x As Long, y As Long
    ReDim Bits(0 To 3, 0 To Width - 1, 0 To Height - 1)
    
    RgbCol LeftTopColor, RS, GS, BS
    RgbCol RightBottomColor, RE, GE, BE
    
    If BlendType = RGBBlend Then
        If LeftToRight Then
            For x = 0 To Width - 1
                rc = (1& * RS - RE) * ((Width - 1 - x) / (Width - 1)) + RE
                GC = (1& * GS - GE) * ((Width - 1 - x) / (Width - 1)) + GE
                BC = (1& * BS - BE) * ((Width - 1 - x) / (Width - 1)) + BE
                For y = 0 To Height - 1
                    Bits(2, x, y) = rc
                    Bits(1, x, y) = GC
                    Bits(0, x, y) = BC
                Next
            Next
        Else
            For y = 0 To Height - 1
                rc = (1& * RS - RE) * ((Height - 1 - y) / (Height - 1)) + RE
                GC = (1& * GS - GE) * ((Height - 1 - y) / (Height - 1)) + GE
                BC = (1& * BS - BE) * ((Height - 1 - y) / (Height - 1)) + BE
                For x = 0 To Width - 1
                    Bits(2, x, y) = rc
                    Bits(1, x, y) = GC
                    Bits(0, x, y) = BC
                Next
            Next
        End If
    ElseIf BlendType = HSLBlend Then
        RGBToHSL RS, GS, BS, HS, SS, LS
        RGBToHSL RE, GE, BE, HE, SE, LE
        If LeftToRight Then
            For x = 0 To Width - 1
                HSLToRGB (1& * HS - HE) * ((Width - 1 - x) / (Width - 1)) + HE, _
                        (1& * SS - SE) * ((Width - 1 - x) / (Width - 1)) + SE, _
                        (1& * LS - LE) * ((Width - 1 - x) / (Width - 1)) + LE, _
                        rc, GC, BC
                For y = 0 To Height - 1
                    Bits(2, x, y) = rc
                    Bits(1, x, y) = GC
                    Bits(0, x, y) = BC
                Next
            Next
        Else
            For y = 0 To Height - 1
                HSLToRGB (1& * HS - HE) * ((Height - 1 - y) / (Height - 1)) + HE, _
                        (1& * SS - SE) * ((Height - 1 - y) / (Height - 1)) + SE, _
                        (1& * LS - LE) * ((Height - 1 - y) / (Height - 1)) + LE, _
                        rc, GC, BC
                For x = 0 To Width - 1
                    Bits(2, x, y) = rc
                    Bits(1, x, y) = GC
                    Bits(0, x, y) = BC
                Next
            Next
        End If
    End If

    Dim BI As BITMAPINFO
    With BI.bmiHeader
        .biSize = Len(BI.bmiHeader)
        .biWidth = Width
        .biHeight = -Height
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = BI_RGB
        .biSizeImage = ((((.biWidth * .biBitCount) + 31) \ 32) * 4) * Abs(.biHeight)
    End With
    hBmp = CreateBitmap(Width, Height, 1&, 32&, ByVal 0)
    SetDIBits 0&, hBmp, 0, Abs(BI.bmiHeader.biHeight), Bits(0, 0, 0), BI, DIB_RGB_COLORS

    Dim IGuid As Guid, PicDst As PictDesc
    With IGuid
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    With PicDst
        .cbSizeofStruct = Len(PicDst)
        .hImage = hBmp
        .picType = vbPicTypeBitmap
    End With
    OleCreatePictureIndirect PicDst, IGuid, True, CreateGradient
End Function

'Helper Functions
Private Sub RgbCol(Col As Long, ByRef R As Byte, ByRef G As Byte, ByRef B As Byte)
    R = Col And &HFF&
    G = (Col And &HFF00&) \ &H100&
    B = (Col And &HFF0000) \ &H10000
End Sub

Private Sub RGBToHSL(ByVal R As Byte, ByVal G As Byte, ByVal B As Byte, H As Single, S As Single, L As Single)
    'http://www.vbAccelerator.com
    Dim Max As Single
    Dim Min As Single
    Dim delta As Single
    Dim rR As Single, rG As Single, rB As Single

    rR = R / 255: rG = G / 255: rB = B / 255

    '{Given: rgb each in [0,1].
    ' Desired: h in [0,360] and s in [0,1], except if s=0, then h=UNDEFINED.}
    Max = Maximum(rR, rG, rB)
    Min = Minimum(rR, rG, rB)
    L = (Max + Min) / 2    '{This is the lightness}
    '{Next calculate saturation}
    If Max = Min Then
        'begin {Acrhomatic case}
        S = 0
        H = 0
        'end {Acrhomatic case}
    Else
        'begin {Chromatic case}
             '{First calculate the saturation.}
        If L <= 0.5 Then
            S = (Max - Min) / (Max + Min)
        Else
            S = (Max - Min) / (2 - Max - Min)
        End If
        
        '{Next calculate the hue.}
        delta = Max - Min
        If rR = Max Then
            H = (rG - rB) / delta    '{Resulting color is between yellow and magenta}
        ElseIf rG = Max Then
            H = 2 + (rB - rR) / delta '{Resulting color is between cyan and yellow}
        ElseIf rB = Max Then
            H = 4 + (rR - rG) / delta '{Resulting color is between magenta and cyan}
        End If
        'Debug.Print h
        'h = h * 60
        'If h < 0# Then
        '     h = h + 360            '{Make degrees be nonnegative}
        'End If
    'end {Chromatic Case}
    End If
'end {RGB_to_HLS}
End Sub

Private Sub HSLToRGB(ByVal H As Single, ByVal S As Single, ByVal L As Single, R As Byte, G As Byte, B As Byte)
    'http://www.vbAccelerator.com
    Dim rR As Single, rG As Single, rB As Single
    Dim Min As Single, Max As Single
    
    If S = 0 Then
        ' Achromatic case:
        rR = L: rG = L: rB = L
    Else
        ' Chromatic case:
        ' delta = Max-Min
        If L <= 0.5 Then
            'S = (Max - Min) / (Max + Min)
            ' Get Min value:
            Min = L * (1 - S)
        Else
            'S = (Max - Min) / (2 - Max - Min)
            ' Get Min value:
            Min = L - S * (1 - L)
        End If
        ' Get the Max value:
        Max = 2 * L - Min
       
        ' Now depending on sector we can evaluate the h,l,s:
        If (H < 1) Then
            rR = Max
            If (H < 0) Then
                rG = Min
                rB = rG - H * (Max - Min)
            Else
                rB = Min
                rG = H * (Max - Min) + rB
            End If
        ElseIf (H < 3) Then
            rG = Max
            If (H < 2) Then
                rB = Min
                rR = rB - (H - 2) * (Max - Min)
            Else
                rR = Min
                rB = (H - 2) * (Max - Min) + rR
            End If
        Else
            rB = Max
            If (H < 4) Then
                rR = Min
                rG = rR - (H - 4) * (Max - Min)
            Else
                rG = Min
                rR = (H - 4) * (Max - Min) + rG
            End If
        End If
    End If
    R = rR * 255: G = rG * 255: B = rB * 255
End Sub

Private Function Maximum(rR As Single, rG As Single, rB As Single) As Single
     'http://www.vbAccelerator.com
    If (rR > rG) Then
        If (rR > rB) Then
            Maximum = rR
        Else
            Maximum = rB
        End If
    Else
        If (rB > rG) Then
            Maximum = rB
        Else
            Maximum = rG
        End If
    End If
End Function

Private Function Minimum(rR As Single, rG As Single, rB As Single) As Single
     'http://www.vbAccelerator.com
    If (rR < rG) Then
        If (rR < rB) Then
            Minimum = rR
        Else
            Minimum = rB
        End If
    Else
        If (rB < rG) Then
            Minimum = rB
        Else
            Minimum = rG
        End If
    End If
End Function

Private Sub AddIconToTray()
  Data.Size = Len(Data)
  Data.Handle = hWnd
  Data.ID = vbNull
  Data.Flags = IconFlag Or TipFlag Or MessageFlag
  Data.CallBackMessage = WM_MOUSEMOVE
  Data.Icon = Icon
  Data.Tip = sFormCaption & vbNullChar
  Call Shell_NotifyIcon(AddIcon, Data)
End Sub

Private Sub DeleteIconFromTray()
  Call Shell_NotifyIcon(DeleteIcon, Data)
End Sub

Public Function MakeSemiTransparent(ByVal hWnd As Long, ByVal Perc As Integer) As Long
    Dim Msg As Long
    On Error Resume Next
     
    Perc = ((100 - Perc) / 100) * 255
    If Perc < 0 Or Perc > 255 Then
        MakeSemiTransparent = 1
    Else
        Msg = GetWindowLong(hWnd, G)
        Msg = Msg Or WS_EX_LAYERED
        SetWindowLong hWnd, G, Msg
        
        ' Set the Form header bottom colour
        Col = myForm.BackColor
        lBottomR = (Col And &HFF&)
        lBottomG = (Col And &HFF00&) / &H100
        lBottomB = (Col And &HFF0000) / &H10000
        
        SetLayeredWindowAttributes hWnd, RGB(lBottomR, lBottomG, lBottomB), Perc, LWA_ALPHA
        MakeSemiTransparent = 0
    End If
    If Err Then
        MakeSemiTransparent = 2
    End If
End Function

Private Function MakeTransparent(ByRef Frm As Form, ByVal TrnsColor As Long)
    Frm.BorderStyle = 0
     
    Dim ScaleSize As Long
    Dim Width, Height As Long 'Width and height of the image on our form
    Dim rgnMain As Long 'The main region which will be skinned then will be applied to our form
    Dim x, y As Long 'Variables containing current X, Y in loop below
    Dim rgnPixel As Long 'A single pixel to be cut out of our image
    Dim rgbColor As Long 'A variable to store a color in the loop below
    Dim dcMain As Long 'The temporary DC of where all the skinning takes place
    Dim bmpMain As Long '1x1 bitmap created when dcMain is created
    
    ScaleSize = Frm.ScaleMode
    Frm.ScaleMode = 3 'Set the scale mode to pixels
    
    'This will get the height and width of the image on our form
    Width = Frm.ScaleX(Frm.Width, vbTwips, vbPixels)
    Height = Frm.ScaleY(Frm.Height, vbTwips, vbPixels)
    'vbHimetric
'    Frm.Width = Width * Screen.TwipsPerPixelX
'    Frm.Height = Height * Screen.TwipsPerPixelY
    
    'This will create our basic region to fit the dimensions of our
    'forms image
    rgnMain = CreateRectRgn(0, 0, Width, Height)
    
    'This will create a DC where all the skinning takes place
    dcMain = CreateCompatibleDC(Frm.hdc)
    bmpMain = SelectObject(dcMain, Frm.Picture.Handle)
    
    For y = 0 To 4
        For x = 0 To 4 'Width
            rgbColor = GetPixel(dcMain, x, y) 'Gets the color of a pixel on dcMain
            
            If rgbColor = TrnsColor Then 'If we found a mask color then cut it out of dcMain
                rgnPixel = CreateRectRgn(x, y, x + 1, y + 1) 'Create a region of a single pixel
                CombineRgn rgnMain, rgnMain, rgnPixel, RGN_XOR 'Cut it out
                DeleteObject rgnPixel 'Delete it from the memory
            End If
        Next x
    Next y
    
    For y = 0 To 4
        For x = (Width - 5) To Width 'Width
            rgbColor = GetPixel(dcMain, x, y) 'Gets the color of a pixel on dcMain
            
            If rgbColor = TrnsColor Then 'If we found a mask color then cut it out of dcMain
                rgnPixel = CreateRectRgn(x, y, x + 1, y + 1) 'Create a region of a single pixel
                CombineRgn rgnMain, rgnMain, rgnPixel, RGN_XOR 'Cut it out
                DeleteObject rgnPixel 'Delete it from the memory
            End If
        Next x
    Next y
     
    'Clear up our memory
    SelectObject dcMain, bmpMain
    DeleteDC dcMain
    DeleteObject bmpMain
    
    If rgnMain <> 0 Then
        SetWindowRgn Frm.hWnd, rgnMain, True 'Apply rgnMain to our form
    End If
     
    Frm.ScaleMode = ScaleSize
End Function

Private Function RemoveTransparent(ByRef Frm As Form)

    Dim Width, Height As Long
    Dim rgnMain As Long
    
    'Get size of form
    Width = Frm.ScaleWidth
    Height = Frm.ScaleHeight
    
    rgnMain = CreateRectRgn(0, 0, Width, Height) 'Create a plain old region
    SetWindowRgn Frm.hWnd, rgnMain, True 'Apply to our window
 
End Function

Private Sub ShowInTheTaskbar(hWnd As Long, bShow As Boolean)
    Dim lStyle As Long
    ShowWindow hWnd, SW_HIDE
    
    lStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
    
    If bShow = False Then
        If lStyle And WS_EX_APPWINDOW Then
            lStyle = lStyle - WS_EX_APPWINDOW
        End If
    Else
        lStyle = lStyle Or WS_EX_APPWINDOW
    End If
    
    SetWindowLong hWnd, GWL_EXSTYLE, lStyle
    
    App.TaskVisible = bShow
    
    ShowWindow hWnd, SW_NORMAL
End Sub

Public Property Get Caption() As String
    Caption = sFormCaption
End Property

Public Property Let Caption(ByVal New_TheCaption As String)
    sFormCaption = New_TheCaption
    PropertyChanged "Caption"
    bPaintForm = False
    Call picForm_Paint
End Property

Public Property Get DisplayIcon() As Boolean
    DisplayIcon = bDisplayIcon
End Property

Public Property Let DisplayIcon(ByVal New_DisplayIcon As Boolean)
    bDisplayIcon = New_DisplayIcon
    PropertyChanged "DisplayIcon"
    Call UserControl_Resize
End Property

Public Property Get Font() As Font
    Set Font = picForm.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set picForm.Font = New_Font
    bFontBold = picForm.FontBold
    bFontItalic = picForm.FontItalic
    dFontSize = picForm.FontSize
    bFontStrikeThru = picForm.FontStrikethru
    bFontUnderline = picForm.FontUnderline
    Call UserControl_Resize
    PropertyChanged "Font"
    bPaintForm = False
    Call picForm_Paint
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = picForm.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    picForm.ForeColor() = New_ForeColor
    lFormCaptionColor = New_ForeColor
    PropertyChanged "ForeColor"
    bPaintForm = False
    Call picForm_Paint
End Property

Public Property Get Icon() As Picture
    Set Icon = imgFormPic.Picture
End Property

Public Property Set Icon(ByVal New_Icon As Picture)
    Set imgFormPic.Picture = New_Icon
    Set myForm.Icon = New_Icon
    PropertyChanged "Icon"
    bDisplayIcon = True
    Call UserControl_Resize
End Property

Public Property Get MaxHeight() As Long
    MaxHeight = lFormMaxHeight
End Property

Public Property Let MaxHeight(ByVal New_MaxHeight As Long)
    lFormMaxHeight = New_MaxHeight
    PropertyChanged "MaxHeight"
End Property

Public Property Get MaxWidth() As Long
    MaxWidth = lFormMaxWidth
End Property

Public Property Let MaxWidth(ByVal New_MaxWidth As Long)
    lFormMaxWidth = New_MaxWidth
    PropertyChanged "MaxWidth"
End Property

Public Property Get MinHeight() As Long
    MinHeight = lFormMinHeight
End Property

Public Property Let MinHeight(ByVal New_MinHeight As Long)
    lFormMinHeight = New_MinHeight
    PropertyChanged "MinHeight"
End Property

Public Property Get MinWidth() As Long
    MinWidth = lFormMinWidth
End Property

Public Property Let MinWidth(ByVal New_MinWidth As Long)
    lFormMinWidth = New_MinWidth
    PropertyChanged "MinWidth"
End Property

Public Property Get ShowCloseButton() As Boolean
    ShowCloseButton = bCloseButton
End Property

Public Property Let ShowCloseButton(ByVal New_ShowCloseButton As Boolean)
    bCloseButton = New_ShowCloseButton
    PropertyChanged "ShowCloseButton"
    Call UserControl_Resize
End Property

Public Property Get ShowMinimiseButton() As Boolean
    ShowMinimiseButton = bMinimiseButton
End Property

Public Property Let ShowMinimiseButton(ByVal New_ShowMinimiseButton As Boolean)
    bMinimiseButton = New_ShowMinimiseButton
    PropertyChanged "ShowMinimiseButton"
    Call UserControl_Resize
End Property

Public Property Get ShowMaximiseButton() As Boolean
    ShowMaximiseButton = bMaximiseButton
End Property

Public Property Let ShowMaximiseButton(ByVal New_ShowMaximiseButton As Boolean)
    bMaximiseButton = New_ShowMaximiseButton
    PropertyChanged "ShowMaximiseButton"
    Call UserControl_Resize
End Property

Public Property Get ShowSystemTrayIcon() As Boolean
    ShowSystemTrayIcon = bSystemTray
End Property

Public Property Let ShowSystemTrayIcon(ByVal New_ShowSystemTrayIcon As Boolean)
    bSystemTray = New_ShowSystemTrayIcon
    PropertyChanged "ShowSystemTrayIcon"
    Call UserControl_Resize
End Property

Public Property Get Style() As xVistaStyles
    Style = xVisualStyles
End Property

Public Property Let Style(val As xVistaStyles)
    ' Determine which color scheme has been selected
    xVisualStyles = val

    ' Set the colour scheme
    Call SelectColorScheme

    picForm.ForeColor() = lFormCaptionColor
    lFormCaptionColor = lFormCaptionColor
    PropertyChanged "Style"

    ' Repaint the control
    Call UserControl_Paint

    ' Draw the Form Header and Buttons
    bPaintForm = False
    Call picForm_Paint
End Property

Public Property Get Transparency() As Boolean
    Transparency = bTransparency
End Property

Public Property Let Transparency(ByVal New_Transparency As Boolean)
    bTransparency = New_Transparency
    PropertyChanged "Transparency"
    Call UserControl_Resize
End Property

Public Property Get TransparencyLevel() As String
    TransparencyLevel = iTransparency
End Property

Public Property Let TransparencyLevel(ByVal New_TransparencyLevel As String)
    iTransparency = New_TransparencyLevel
    PropertyChanged "TransparencyLevel"
    Call UserControl_Resize
End Property

Private Sub SelectColorScheme()
    Select Case xVisualStyles
        Case VistaLite
            lFormCaptionColor = &H0&
            lFormGradientBottom = &HEAD1B9
            lFormGradientTop = &HD0B498
            lFormInnerBorder = &HE7D3C1
            lFormMiddleBorder = &HE4CF28
            lFormOuterBorder = &H0&
            lButtonGradientBottom(1) = &HCCB198
            lButtonGradientBottom(2) = &HD1B79E
            lButtonGradientBottom(3) = &HD8BEA4
            lButtonGradientBottom(4) = &HDFC5AC
            lButtonGradientBottom(5) = &HDFC5AC
            lButtonGradientBottom(6) = &HE5CBB2
            lButtonGradientBottom(7) = &HE9D0B7
            lButtonGradientBottomClicked(1) = &H523B20
            lButtonGradientBottomClicked(2) = &H5B441F
            lButtonGradientBottomClicked(3) = &H736223
            lButtonGradientBottomClicked(4) = &H918727
            lButtonGradientBottomClicked(5) = &H918727
            lButtonGradientBottomClicked(6) = &HACA82B
            lButtonGradientBottomClicked(7) = &HC8C927
            lButtonGradientBottomHover(1) = &HA3732D
            lButtonGradientBottomHover(2) = &HAF7B2C
            lButtonGradientBottomHover(3) = &HBF892C
            lButtonGradientBottomHover(4) = &HD09A2C
            lButtonGradientBottomHover(5) = &HD09A2C
            lButtonGradientBottomHover(6) = &HDFA929
            lButtonGradientBottomHover(7) = &HEBC624
            lButtonGradientTop = &HE7D3C1
            lButtonGradientTopClicked = &H9C886E
            lButtonGradientTopHover = &HEFCB96
            lButtonInnerBorder = &HF2E7DE
            lButtonOuterBorder = &H886F5D
            lCloseButtonGradientBottom(1) = &H2C43B8
            lCloseButtonGradientBottom(2) = &H3249BA
            lCloseButtonGradientBottom(3) = &H3F54BF
            lCloseButtonGradientBottom(4) = &H4F62C5
            lCloseButtonGradientBottom(5) = &H4F62C5
            lCloseButtonGradientBottom(6) = &H6373CD
            lCloseButtonGradientBottom(7) = &H7685D5
            lCloseButtonGradientBottomClicked(1) = &H1883&
            lCloseButtonGradientBottomClicked(2) = &H1987&
            lCloseButtonGradientBottomClicked(3) = &H12B85
            lCloseButtonGradientBottomClicked(4) = &H124391
            lCloseButtonGradientBottomClicked(5) = &H124391
            lCloseButtonGradientBottomClicked(6) = &H2C68A8
            lCloseButtonGradientBottomClicked(7) = &H4A93C1
            lCloseButtonGradientBottomHover(1) = &H223D2
            lCloseButtonGradientBottomHover(2) = &H223D2
            lCloseButtonGradientBottomHover(3) = &HD33D5
            lCloseButtonGradientBottomHover(4) = &H2151DA
            lCloseButtonGradientBottomHover(5) = &H2151DA
            lCloseButtonGradientBottomHover(6) = &H3974E0
            lCloseButtonGradientBottomHover(7) = &H56A0E8
            lCloseButtonGradientTop = &H929FE4
            lCloseButtonGradientTopClicked = &H768FBF
            lCloseButtonGradientTopHover = &HADB9FC
            lCloseButtonInnerBorder = &HCCD3F4
            lCloseButtonOuterBorder = &H221443
        Case VistaDark
            lFormCaptionColor = &HFFFFFF
            lFormGradientBottom = &H322624
            lFormGradientTop = &H9E9794
            lFormInnerBorder = &HE0E0E0
            lFormMiddleBorder = &HB9B8B4
            lFormOuterBorder = &H0&
            lButtonGradientBottom(1) = &H433E35
            lButtonGradientBottom(2) = &H464138
            lButtonGradientBottom(3) = &H413E36
            lButtonGradientBottom(4) = &H403D35
            lButtonGradientBottom(5) = &H403C37
            lButtonGradientBottom(6) = &H413D38
            lButtonGradientBottom(7) = &H423E39
            lButtonGradientBottomClicked(1) = &H2B2B2B
            lButtonGradientBottomClicked(2) = &H353535
            lButtonGradientBottomClicked(3) = &H414141
            lButtonGradientBottomClicked(4) = &H4F4F4F
            lButtonGradientBottomClicked(5) = &H4F4F4F
            lButtonGradientBottomClicked(6) = &H535353
            lButtonGradientBottomClicked(7) = &H666666
            lButtonGradientBottomHover(1) = &H7B695B
            lButtonGradientBottomHover(2) = &H816F60
            lButtonGradientBottomHover(3) = &H8C7B6C
            lButtonGradientBottomHover(4) = &H9A8979
            lButtonGradientBottomHover(5) = &H9A8979
            lButtonGradientBottomHover(6) = &HAA9A88
            lButtonGradientBottomHover(7) = &HBAAA97
            lButtonGradientTop = &HC0BDB8
            lButtonGradientTopClicked = &H4D4D4D
            lButtonGradientTopHover = &HD1CAC4
            lButtonInnerBorder = &HB9B8B4
            lButtonOuterBorder = &H322624
            lCloseButtonGradientBottom(1) = &H2C43B8
            lCloseButtonGradientBottom(2) = &H3249BA
            lCloseButtonGradientBottom(3) = &H3F54BF
            lCloseButtonGradientBottom(4) = &H4F62C5
            lCloseButtonGradientBottom(5) = &H4F62C5
            lCloseButtonGradientBottom(6) = &H6373CD
            lCloseButtonGradientBottom(7) = &H7685D5
            lCloseButtonGradientBottomClicked(1) = &H1883&
            lCloseButtonGradientBottomClicked(2) = &H1987&
            lCloseButtonGradientBottomClicked(3) = &H12B85
            lCloseButtonGradientBottomClicked(4) = &H124391
            lCloseButtonGradientBottomClicked(5) = &H124391
            lCloseButtonGradientBottomClicked(6) = &H2C68A8
            lCloseButtonGradientBottomClicked(7) = &H4A93C1
            lCloseButtonGradientBottomHover(1) = &H223D2
            lCloseButtonGradientBottomHover(2) = &H223D2
            lCloseButtonGradientBottomHover(3) = &HD33D5
            lCloseButtonGradientBottomHover(4) = &H2151DA
            lCloseButtonGradientBottomHover(5) = &H2151DA
            lCloseButtonGradientBottomHover(6) = &H3974E0
            lCloseButtonGradientBottomHover(7) = &H56A0E8
            lCloseButtonGradientTop = &H929FE4
            lCloseButtonGradientTopClicked = &H768FBF
            lCloseButtonGradientTopHover = &HADB9FC
            lCloseButtonInnerBorder = &HCCD3F4
            lCloseButtonOuterBorder = &H221443
    End Select
End Sub

Private Sub UserControlsCreate()
    If iNumControls = 0 Then
        ' Create the controls only once
        iNumControls = 1
        
        ' Add the Form Header picturebox
        Set picForm = UserControl.Controls.Add("VB.PictureBox", "picForm")
        picForm.AutoRedraw = True
        picForm.BorderStyle = 0
        picForm.Visible = True
        
        ' Add the Form Timer
        Set tmrMouseMove = UserControl.Controls.Add("VB.Timer", "tmrMouseMove")
        tmrMouseMove.Enabled = False
        tmrMouseMove.Interval = 10
               
        ' Add the Form Header Image
        Set imgFormPic = Controls.Add("VB.Image", "imgFormPic", picForm)
        Set imgFormPic.Picture = Nothing
    End If
End Sub

Private Sub moveForm_Load()
    If bSystemTray = True Then
        Call AddIconToTray
    End If

    If lFormMaxHeight > 0 Or lFormMaxWidth > 0 Then
        If lFormMaxHeight > 0 Then myForm.Height = lFormMaxHeight
        If lFormMaxWidth > 0 Then myForm.Width = lFormMaxWidth
    Else
        If myForm.WindowState = vbMaximized Then
            myForm.WindowState = vbMaximized
        Else
            myForm.WindowState = vbNormal
        End If
    End If

    Call UserControl_Paint
End Sub

Private Sub moveForm_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Dim lngReturnValue As Long
    
    If Button = vbLeftButton Then
        bMouseClicked = True
        If bMinimiseButtonHover = True Then
            bCloseButtonHover = False
            bCloseButtonClicked = False
            bMaximiseButtonClicked = False
            bMaximiseButtonHover = False
            bMinimiseButtonClicked = True
            Call DrawMinimiseButton
            picForm.Refresh
        ElseIf bMaximiseButtonHover = True Then
            bCloseButtonHover = False
            bCloseButtonClicked = False
            bMaximiseButtonClicked = True
            bMinimiseButtonClicked = False
            bMinimiseButtonHover = False
            Call DrawMaximiseButton
            picForm.Refresh
        ElseIf bCloseButtonHover = True Then
            bCloseButtonClicked = True
            bMaximiseButtonClicked = False
            bMaximiseButtonHover = False
            bMinimiseButtonClicked = False
            bMinimiseButtonHover = False
            Call DrawCloseButton
            picForm.Refresh
        End If
        
        If x < (UserControl.Width - 1500) And (y > 30 And y <= 375) Then
            Call ReleaseCapture
            myForm.MousePointer = vbSizeAll
            lngReturnValue = SendMessage(moveForm.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
            myForm.MousePointer = vbNormal
        End If
        
        Dim RLeft As Single
        Dim RTop As Single
        Dim RRight As Boolean
        Dim RBottom As Boolean
        Dim StartX As Single
        Dim StartY As Single
        Dim MX As Single
        Dim MY As Single

        ' Read the mouse pointer screen position on the beginning
        GetMouseXY StartX, StartY
        ' We don't use the X,Y arguments which are incorect when MouseDown event called
        ' from other screen objects (like Label1 and Image1 hereunder)
        x = StartX - myForm.Left
        y = StartY - myForm.Top

        ' Flags indicating "from where" the form is being resized
        RLeft = IIf(x < BorderWidth, myForm.Width, 0)
        RTop = IIf(y < BorderHeight, myForm.Height, 0)
        RRight = (x > myForm.Width - BorderWidth)
        RBottom = (y > myForm.Height - BorderHeight)

        ' Place the mouse pointer on the form border for more accuracy
        If RLeft Then SetMouseXY myForm.Left, myForm.Top + y
        If RTop Then SetMouseXY myForm.Left + x, myForm.Top
        If RRight Then SetMouseXY myForm.Left + myForm.Width, myForm.Top + y
        If RBottom Then SetMouseXY myForm.Left + x, myForm.Top + myForm.Height

        ' Save the mouse pointer screen position on the beginning in variables
        GetMouseXY StartX, StartY
        ' While left mouse button is pressed
        While GetAsyncKeyState(vbLeftButton) < 0
            ' Read the actual mouse pointer screen position
            GetMouseXY MX, MY

            If RRight Or RLeft Or RBottom Or RTop Then      ' If the form is resized (not moved)
                If lFormMaxHeight > 0 And myForm.Height - 5 > lFormMaxHeight Then
                    myForm.Height = lFormMaxHeight
                    Exit Sub
                End If
                If lFormMaxWidth > 0 And myForm.Width - 5 > lFormMaxWidth Then
                    myForm.Width = lFormMaxWidth
                    Exit Sub
                End If
                If lFormMinHeight > 0 And myForm.Height + 5 < lFormMinHeight Then
                    myForm.Height = lFormMinHeight
                    Exit Sub
                End If
                If lFormMinWidth > 0 And myForm.Width + 5 < lFormMinWidth Then
                    myForm.Width = lFormMinWidth
                    Exit Sub
                End If
                If RLeft And RLeft + StartX - MX > BorderWidth * 2 Then myForm.Move MX, myForm.Top, RLeft + StartX - MX
                If RTop And RTop + StartY - MY > BorderHeight * 2 Then myForm.Move myForm.Left, MY, myForm.Width, RTop + StartY - MY
                If RRight And MX - myForm.Left > BorderWidth * 2 Then myForm.Width = MX - myForm.Left
                If RBottom And MY - myForm.Top > BorderHeight * 2 Then myForm.Height = MY - myForm.Top
''                Else                                            ' If the form is moved (not resized)
''                    MousePointer = vbSizeAll                    ' Sets the mouse cursor showing move
''                    myForm.Move MX - X, MY - Y                  ' Actually moves the form on screen
            End If
            DoEvents                                        ' To allow Windows painting events
        Wend
        
        If bUnloadForm = False Then
            Call DrawMinimiseButton
            Call DrawMaximiseButton
            Call DrawCloseButton
            myForm.Refresh
        End If
    End If
End Sub

Private Sub moveForm_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Reset the mouse cursor when left mouse button is not pressed
    If GetAsyncKeyState(vbLeftButton) >= 0 Then myForm.MousePointer = 0
    
    ' Set the correct mouse cursor according to its position on the form
    If (x + BorderWidth) > myForm.Width - BorderWidth Or x < BorderWidth Then myForm.MousePointer = vbSizeWE
    If bRightClick = False And (y > myForm.Height - BorderHeight Or y < 30) Then myForm.MousePointer = vbSizeNS
    If y > 30 And y <= (UserControl.Height + 15) And (x > 30 And x < UserControl.Width - 30) Then myForm.MousePointer = vbNormal
    If (x + BorderWidth > myForm.Width - BorderWidth And y > myForm.Height - BorderHeight) Or (x < BorderWidth And y < BorderHeight) Then myForm.MousePointer = vbSizeNWSE
    If (x + BorderWidth > myForm.Width - BorderWidth And y < BorderHeight) Or (x < BorderWidth And y > myForm.Height - BorderHeight) Then myForm.MousePointer = vbSizeNESW
End Sub

Private Sub moveForm_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If bMinimiseButtonHover = True And bMinimiseButton = True And bMinimiseButtonClicked = True Then
        myForm.WindowState = 1
    ElseIf bMaximiseButtonHover = True And bMaximiseButton = True And bMaximiseButtonClicked = True Then
        If myForm.WindowState <> vbMaximized Then
            If lFormMaxHeight > 0 Or lFormMaxWidth > 0 Then
                If lFormMaxHeight > 0 Then myForm.Height = lFormMaxHeight
                If lFormMaxWidth > 0 Then myForm.Width = lFormMaxWidth
            Else
                myForm.WindowState = vbMaximized
            End If
        Else
            myForm.WindowState = vbNormal
        End If
    ElseIf bCloseButtonHover = True And bCloseButton = True And bCloseButtonClicked = True Then
        tmrMouseMove.Enabled = False
        bUnloadForm = True
        Unload moveForm
        Unload myForm
    Else
        bCloseButtonClicked = False
        bMaximiseButtonClicked = False
        bMinimiseButtonClicked = False
        Call DrawMinimiseButton
        Call DrawMaximiseButton
        Call DrawCloseButton
        picForm.Refresh
    End If
    bMouseClicked = False
    bMouseOnForm = False
End Sub

' Reads actual mouse pointer screen position and convert it to TWIP scale
Private Sub GetMouseXY(x As Single, y As Single)
    Dim lpPoint As POINTAPI

    GetCursorPos lpPoint
    x = lpPoint.x * TwipX
    y = lpPoint.y * TwipY
End Sub

' Places mouse pointer on given screen position given in TWIP scale
Private Sub SetMouseXY(ByVal x As Single, ByVal y As Single)
    SetCursorPos x / TwipX, y / TwipY
End Sub

Private Sub moveForm_Resize()
    Call UserControl_Resize
    myForm.Refresh
End Sub

Private Sub moveForm_Terminate()
    Call DeleteIconFromTray
End Sub

Private Sub moveForm_Unload(Cancel As Integer)
    Call DeleteIconFromTray
  
    ' Destroy the menu before exiting
    If lSysTrayMenu Then
      Call DestroyMenu(lSysTrayMenu)
      bUnloadForm = True
    End If
End Sub

Private Sub picForm_DblClick()
    If myForm.WindowState <> vbMaximized And bMaximiseButton = True Then
        If lFormMaxHeight > 0 Or lFormMaxWidth > 0 Then
            If lFormMaxHeight > 0 Then myForm.Height = lFormMaxHeight
            If lFormMaxWidth > 0 Then myForm.Width = lFormMaxWidth
        Else
            myForm.WindowState = vbMaximized
        End If
    Else
        myForm.WindowState = vbNormal
    End If
End Sub

Private Sub picForm_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call moveForm_MouseDown(Button, Shift, x, y)
End Sub

Private Sub picForm_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Set the mouse to default
    If y > 30 And (x > 30 Or x < UserControl.Width - 30) Then myForm.MousePointer = vbNormal
    
    ' Determine which button has hover focus
    If (x >= (UserControl.Width - 1470) And x <= (UserControl.Width - 1050)) And (y >= 75 And y <= 300) Then
        bCloseButtonHover = False
        bMaximiseButtonHover = False
        bMinimiseButtonHover = True
        If Button = 1 Then
            If bMouseClicked = False Then
                bMinimiseButtonClicked = True
            End If
        End If
    ElseIf x >= (UserControl.Width - 1005) And x <= (UserControl.Width - 575) And (y >= 75 And y <= 300) Then
        bCloseButtonHover = False
        bMaximiseButtonHover = True
        bMinimiseButtonHover = False
        If Button = 1 Then
            If bMouseClicked = False Then
                bMaximiseButtonClicked = True
            End If
        End If
    ElseIf x >= (UserControl.Width - 520) And x <= (UserControl.Width - 105) And (y >= 75 And y <= 300) Then
        bCloseButtonHover = True
        bMaximiseButtonHover = False
        bMinimiseButtonHover = False
        If Button = 1 Then
            If bMouseClicked = False Then
                bCloseButtonClicked = True
            End If
        End If
    Else
        bCloseButtonClicked = False
        bCloseButtonHover = False
        bMaximiseButtonClicked = False
        bMaximiseButtonHover = False
        bMinimiseButtonClicked = False
        bMinimiseButtonHover = False
    End If
    
    ' Enable the timer
    tmrMouseMove.Enabled = True
    
    Call moveForm_MouseMove(Button, Shift, x, y)
End Sub

Private Sub DrawCloseButton()
    On Error GoTo Errhandler
    If bCloseButton = True Then
        ' Draw the Buttons Outside Border
        picForm.Line (UserControl.Width - 510, UserControl.Extender.Top + 75)-(UserControl.Width - 105, UserControl.Extender.Top + 75), lCloseButtonOuterBorder
        picForm.Line (UserControl.Width - 510, UserControl.Extender.Top + 315)-(UserControl.Width - 105, UserControl.Extender.Top + 315), lCloseButtonOuterBorder
        picForm.Line (UserControl.Width - 525, UserControl.Extender.Top + 90)-(UserControl.Width - 525, UserControl.Extender.Top + 315), lCloseButtonOuterBorder
        picForm.Line (UserControl.Width - 105, UserControl.Extender.Top + 90)-(UserControl.Width - 105, UserControl.Extender.Top + 315), lCloseButtonOuterBorder

        ' Draw the Buttons Inside Border
        picForm.Line (UserControl.Width - 495, UserControl.Extender.Top + 90)-(UserControl.Width - 120, UserControl.Extender.Top + 90), lCloseButtonInnerBorder
        picForm.Line (UserControl.Width - 510, UserControl.Extender.Top + 300)-(UserControl.Width - 105, UserControl.Extender.Top + 300), lCloseButtonInnerBorder
        picForm.Line (UserControl.Width - 510, UserControl.Extender.Top + 90)-(UserControl.Width - 510, UserControl.Extender.Top + 300), lCloseButtonInnerBorder
        picForm.Line (UserControl.Width - 120, UserControl.Extender.Top + 90)-(UserControl.Width - 120, UserControl.Extender.Top + 300), lCloseButtonInnerBorder

        ' Close Button Top Gradient Base Colour
        iVertical = 105
        For i = 1 To 6
            If bCloseButtonHover = False And bCloseButtonClicked = False Then
                picForm.Line (UserControl.Width - 495, UserControl.Extender.Top + iVertical)-(UserControl.Width - 120, UserControl.Extender.Top + iVertical), lCloseButtonGradientTop
            ElseIf bCloseButtonHover = True And bCloseButtonClicked = False Then
                picForm.Line (UserControl.Width - 495, UserControl.Extender.Top + iVertical)-(UserControl.Width - 120, UserControl.Extender.Top + iVertical), lCloseButtonGradientTopHover
            ElseIf bCloseButtonHover = True And bCloseButtonClicked = True Then
                picForm.Line (UserControl.Width - 495, UserControl.Extender.Top + iVertical)-(UserControl.Width - 120, UserControl.Extender.Top + iVertical), lCloseButtonGradientTopClicked
            End If
            iVertical = iVertical + 15
        Next
        
        ' Button Bottom Gradient Base Colour
        For i = 1 To 7
            If bCloseButtonHover = False And bCloseButtonClicked = False Then
                picForm.Line (UserControl.Width - 495, UserControl.Extender.Top + iVertical)-(UserControl.Width - 120, UserControl.Extender.Top + iVertical), lCloseButtonGradientBottom(i)
            ElseIf bCloseButtonHover = True And bCloseButtonClicked = False Then
                picForm.Line (UserControl.Width - 495, UserControl.Extender.Top + iVertical)-(UserControl.Width - 120, UserControl.Extender.Top + iVertical), lCloseButtonGradientBottomHover(i)
            ElseIf bCloseButtonHover = True And bCloseButtonClicked = True Then
                picForm.Line (UserControl.Width - 495, UserControl.Extender.Top + iVertical)-(UserControl.Width - 120, UserControl.Extender.Top + iVertical), lCloseButtonGradientBottomClicked(i)
            End If
            iVertical = iVertical + 15
        Next
        
        ' Draw the Close Button Display
        ' Outside Borders
        picForm.Line (UserControl.Width - 375, UserControl.Extender.Top + 135)-(UserControl.Width - 330, UserControl.Extender.Top + 135), &H665653
        picForm.Line (UserControl.Width - 285, UserControl.Extender.Top + 135)-(UserControl.Width - 240, UserControl.Extender.Top + 135), &H665653
        
        picForm.Line (UserControl.Width - 390, UserControl.Extender.Top + 150)-(UserControl.Width - 375, UserControl.Extender.Top + 150), &H665653
        picForm.Line (UserControl.Width - 330, UserControl.Extender.Top + 150)-(UserControl.Width - 315, UserControl.Extender.Top + 150), &H665653
        picForm.Line (UserControl.Width - 300, UserControl.Extender.Top + 150)-(UserControl.Width - 285, UserControl.Extender.Top + 150), &H665653
        picForm.Line (UserControl.Width - 240, UserControl.Extender.Top + 150)-(UserControl.Width - 225, UserControl.Extender.Top + 150), &H665653
        
        picForm.Line (UserControl.Width - 375, UserControl.Extender.Top + 165)-(UserControl.Width - 360, UserControl.Extender.Top + 165), &H665653
        picForm.Line (UserControl.Width - 315, UserControl.Extender.Top + 165)-(UserControl.Width - 300, UserControl.Extender.Top + 165), &H665653
        picForm.Line (UserControl.Width - 255, UserControl.Extender.Top + 165)-(UserControl.Width - 240, UserControl.Extender.Top + 165), &H665653
       
        picForm.Line (UserControl.Width - 360, UserControl.Extender.Top + 180)-(UserControl.Width - 345, UserControl.Extender.Top + 180), &H665653
        picForm.Line (UserControl.Width - 270, UserControl.Extender.Top + 180)-(UserControl.Width - 255, UserControl.Extender.Top + 180), &H665653

        picForm.Line (UserControl.Width - 345, UserControl.Extender.Top + 195)-(UserControl.Width - 330, UserControl.Extender.Top + 195), &H665653
        picForm.Line (UserControl.Width - 285, UserControl.Extender.Top + 195)-(UserControl.Width - 270, UserControl.Extender.Top + 195), &H665653

        picForm.Line (UserControl.Width - 360, UserControl.Extender.Top + 210)-(UserControl.Width - 345, UserControl.Extender.Top + 210), &H665653
        picForm.Line (UserControl.Width - 270, UserControl.Extender.Top + 210)-(UserControl.Width - 255, UserControl.Extender.Top + 210), &H665653

        picForm.Line (UserControl.Width - 375, UserControl.Extender.Top + 225)-(UserControl.Width - 360, UserControl.Extender.Top + 225), &H665653
        picForm.Line (UserControl.Width - 315, UserControl.Extender.Top + 225)-(UserControl.Width - 300, UserControl.Extender.Top + 225), &H665653
        picForm.Line (UserControl.Width - 255, UserControl.Extender.Top + 225)-(UserControl.Width - 240, UserControl.Extender.Top + 225), &H665653

        picForm.Line (UserControl.Width - 390, UserControl.Extender.Top + 240)-(UserControl.Width - 375, UserControl.Extender.Top + 240), &H665653
        picForm.Line (UserControl.Width - 330, UserControl.Extender.Top + 240)-(UserControl.Width - 315, UserControl.Extender.Top + 240), &H665653
        picForm.Line (UserControl.Width - 300, UserControl.Extender.Top + 240)-(UserControl.Width - 285, UserControl.Extender.Top + 240), &H665653
        picForm.Line (UserControl.Width - 240, UserControl.Extender.Top + 240)-(UserControl.Width - 225, UserControl.Extender.Top + 240), &H665653

        picForm.Line (UserControl.Width - 375, UserControl.Extender.Top + 255)-(UserControl.Width - 330, UserControl.Extender.Top + 255), &H665653
        picForm.Line (UserControl.Width - 285, UserControl.Extender.Top + 255)-(UserControl.Width - 240, UserControl.Extender.Top + 255), &H665653

        ' Inside Button Colours
        picForm.Line (UserControl.Width - 375, UserControl.Extender.Top + 150)-(UserControl.Width - 330, UserControl.Extender.Top + 150), &HFFFFFF
        picForm.Line (UserControl.Width - 285, UserControl.Extender.Top + 150)-(UserControl.Width - 240, UserControl.Extender.Top + 150), &HFFFFFF
        
        picForm.Line (UserControl.Width - 360, UserControl.Extender.Top + 165)-(UserControl.Width - 315, UserControl.Extender.Top + 165), &HFFFFFF
        picForm.Line (UserControl.Width - 300, UserControl.Extender.Top + 165)-(UserControl.Width - 255, UserControl.Extender.Top + 165), &HFFFFFF
        
        picForm.Line (UserControl.Width - 345, UserControl.Extender.Top + 180)-(UserControl.Width - 270, UserControl.Extender.Top + 180), &HFFFFFF
        
        picForm.Line (UserControl.Width - 330, UserControl.Extender.Top + 195)-(UserControl.Width - 285, UserControl.Extender.Top + 195), &HE9E9E9
        
        picForm.Line (UserControl.Width - 345, UserControl.Extender.Top + 210)-(UserControl.Width - 270, UserControl.Extender.Top + 210), &HE2E2E2
        
        picForm.Line (UserControl.Width - 360, UserControl.Extender.Top + 225)-(UserControl.Width - 315, UserControl.Extender.Top + 225), &HDCDCDC
        picForm.Line (UserControl.Width - 300, UserControl.Extender.Top + 225)-(UserControl.Width - 255, UserControl.Extender.Top + 225), &HDCDCDC
        
        picForm.Line (UserControl.Width - 375, UserControl.Extender.Top + 240)-(UserControl.Width - 330, UserControl.Extender.Top + 240), &HD7D7D7
        picForm.Line (UserControl.Width - 285, UserControl.Extender.Top + 240)-(UserControl.Width - 240, UserControl.Extender.Top + 240), &HD7D7D7
        
        ' Border Single Pixel Colours
        If xVisualStyles = VistaLite Then
            picForm.Line (UserControl.Width - 525, UserControl.Extender.Top + 75)-(UserControl.Width - 510, UserControl.Extender.Top + 75), &HBCA6A0
            picForm.Line (UserControl.Width - 105, UserControl.Extender.Top + 75)-(UserControl.Width - 90, UserControl.Extender.Top + 75), &HBCA6A0
            picForm.Line (UserControl.Width - 525, UserControl.Extender.Top + 315)-(UserControl.Width - 510, UserControl.Extender.Top + 315), &HBCA6A0
            picForm.Line (UserControl.Width - 105, UserControl.Extender.Top + 315)-(UserControl.Width - 90, UserControl.Extender.Top + 315), &HBCA6A0
        End If
    End If
Errhandler:
End Sub

Private Sub DrawMinimiseButton()
    On Error GoTo Errhandler
    
    ' Draw the Buttons
    If bMinimiseButton = True Then
        ' Draw the Minimise Buttons Outside Border
        If xVisualStyles = VistaLite Then
            picForm.Line (UserControl.Width - 1485, UserControl.Extender.Top + 75)-(UserControl.Width - 1035, UserControl.Extender.Top + 75), &HC4AA91
            picForm.Line (UserControl.Width - 1485, UserControl.Extender.Top + 315)-(UserControl.Width - 1035, UserControl.Extender.Top + 315), &HC4AA91
        End If
        
        picForm.Line (UserControl.Width - 1470, UserControl.Extender.Top + 75)-(UserControl.Width - 1050, UserControl.Extender.Top + 75), lButtonOuterBorder
        picForm.Line (UserControl.Width - 1470, UserControl.Extender.Top + 315)-(UserControl.Width - 1050, UserControl.Extender.Top + 315), lButtonOuterBorder
        picForm.Line (UserControl.Width - 1485, UserControl.Extender.Top + 90)-(UserControl.Width - 1485, UserControl.Extender.Top + 315), lButtonOuterBorder
        picForm.Line (UserControl.Width - 1050, UserControl.Extender.Top + 90)-(UserControl.Width - 1050, UserControl.Extender.Top + 315), lButtonOuterBorder

        ' Draw the Buttons Inside Border
        picForm.Line (UserControl.Width - 1470, UserControl.Extender.Top + 90)-(UserControl.Width - 1070, UserControl.Extender.Top + 90), lButtonInnerBorder
        picForm.Line (UserControl.Width - 1470, UserControl.Extender.Top + 300)-(UserControl.Width - 1055, UserControl.Extender.Top + 300), lButtonInnerBorder
        picForm.Line (UserControl.Width - 1470, UserControl.Extender.Top + 90)-(UserControl.Width - 1470, UserControl.Extender.Top + 300), lButtonInnerBorder
        picForm.Line (UserControl.Width - 1070, UserControl.Extender.Top + 90)-(UserControl.Width - 1070, UserControl.Extender.Top + 300), lButtonInnerBorder

        ' Button Top Gradient Base Colour
        iVertical = 105
        For i = 1 To 6
            If bMinimiseButtonHover = False And bMinimiseButtonClicked = False Then
                picForm.Line (UserControl.Width - 1455, UserControl.Extender.Top + iVertical)-(UserControl.Width - 1070, UserControl.Extender.Top + iVertical), lButtonGradientTop
            ElseIf bMinimiseButtonHover = True And bMinimiseButtonClicked = False Then
                picForm.Line (UserControl.Width - 1455, UserControl.Extender.Top + iVertical)-(UserControl.Width - 1070, UserControl.Extender.Top + iVertical), lButtonGradientTopHover
            ElseIf bMinimiseButtonHover = True And bMinimiseButtonClicked = True Then
                picForm.Line (UserControl.Width - 1455, UserControl.Extender.Top + iVertical)-(UserControl.Width - 1070, UserControl.Extender.Top + iVertical), lButtonGradientTopClicked
            End If
            iVertical = iVertical + 15
        Next
        
        ' Button Bottom Gradient Base Colour
        For i = 1 To 7
            If bMinimiseButtonHover = False And bMinimiseButtonClicked = False Then
                picForm.Line (UserControl.Width - 1455, UserControl.Extender.Top + iVertical)-(UserControl.Width - 1070, UserControl.Extender.Top + iVertical), lButtonGradientBottom(i)
            ElseIf bMinimiseButtonHover = True And bMinimiseButtonClicked = False Then
                picForm.Line (UserControl.Width - 1455, UserControl.Extender.Top + iVertical)-(UserControl.Width - 1070, UserControl.Extender.Top + iVertical), lButtonGradientBottomHover(i)
            ElseIf bMinimiseButtonHover = True And bMinimiseButtonClicked = True Then
                picForm.Line (UserControl.Width - 1455, UserControl.Extender.Top + iVertical)-(UserControl.Width - 1070, UserControl.Extender.Top + iVertical), lButtonGradientBottomClicked(i)
            End If
            iVertical = iVertical + 15
        Next
         
        ' Draw the Minimise Button Display
        ' Outside Borders
        If xVisualStyles = VistaLite Then
            picForm.Line (UserControl.Width - 1350, UserControl.Extender.Top + 195)-(UserControl.Width - 1175, UserControl.Extender.Top + 195), &H9E8A7C    ' Top Border
            picForm.Line (UserControl.Width - 1350, UserControl.Extender.Top + 255)-(UserControl.Width - 1175, UserControl.Extender.Top + 255), &H9E8A7C    ' Bottom Border
        End If
        
        picForm.Line (UserControl.Width - 1335, UserControl.Extender.Top + 195)-(UserControl.Width - 1190, UserControl.Extender.Top + 195), &H665653    ' Top Border
        picForm.Line (UserControl.Width - 1335, UserControl.Extender.Top + 255)-(UserControl.Width - 1190, UserControl.Extender.Top + 255), &H665653    ' Bottom Border
        picForm.Line (UserControl.Width - 1350, UserControl.Extender.Top + 210)-(UserControl.Width - 1350, UserControl.Extender.Top + 255), &H665653    ' Left Border
        picForm.Line (UserControl.Width - 1190, UserControl.Extender.Top + 210)-(UserControl.Width - 1190, UserControl.Extender.Top + 255), &H665653    ' Right Border
    
        ' Inside Button Display
        picForm.Line (UserControl.Width - 1335, UserControl.Extender.Top + 210)-(UserControl.Width - 1190, UserControl.Extender.Top + 210), &HFFFFFF    ' Top Border
        picForm.Line (UserControl.Width - 1335, UserControl.Extender.Top + 225)-(UserControl.Width - 1190, UserControl.Extender.Top + 225), &HDCDCDC    ' Middle Border
        picForm.Line (UserControl.Width - 1335, UserControl.Extender.Top + 240)-(UserControl.Width - 1190, UserControl.Extender.Top + 240), &HD7D7D7    ' Bottom Border
    End If
Errhandler:
End Sub

Private Sub DrawMaximiseButton()
    On Error GoTo Errhandler
    If bMaximiseButton = True Then
        ' Draw the Maximise Buttons Outside Border
        If xVisualStyles = VistaLite Then
            picForm.Line (UserControl.Width - 1005, UserControl.Extender.Top + 75)-(UserControl.Width - 555, UserControl.Extender.Top + 75), &HC4AA91
            picForm.Line (UserControl.Width - 1005, UserControl.Extender.Top + 315)-(UserControl.Width - 555, UserControl.Extender.Top + 315), &HC4AA91
        End If
        
        picForm.Line (UserControl.Width - 990, UserControl.Extender.Top + 75)-(UserControl.Width - 570, UserControl.Extender.Top + 75), lButtonOuterBorder
        picForm.Line (UserControl.Width - 990, UserControl.Extender.Top + 315)-(UserControl.Width - 570, UserControl.Extender.Top + 315), lButtonOuterBorder
        picForm.Line (UserControl.Width - 1005, UserControl.Extender.Top + 90)-(UserControl.Width - 1005, UserControl.Extender.Top + 315), lButtonOuterBorder
        picForm.Line (UserControl.Width - 570, UserControl.Extender.Top + 90)-(UserControl.Width - 570, UserControl.Extender.Top + 315), lButtonOuterBorder

        ' Draw the Buttons Inside Border
        picForm.Line (UserControl.Width - 980, UserControl.Extender.Top + 90)-(UserControl.Width - 590, UserControl.Extender.Top + 90), lButtonInnerBorder
        picForm.Line (UserControl.Width - 990, UserControl.Extender.Top + 300)-(UserControl.Width - 575, UserControl.Extender.Top + 300), lButtonInnerBorder
        picForm.Line (UserControl.Width - 990, UserControl.Extender.Top + 90)-(UserControl.Width - 990, UserControl.Extender.Top + 300), lButtonInnerBorder
        picForm.Line (UserControl.Width - 590, UserControl.Extender.Top + 90)-(UserControl.Width - 590, UserControl.Extender.Top + 300), lButtonInnerBorder

        ' Button Top Gradient Base Colour
        iVertical = 105
        For i = 1 To 6
            If bMaximiseButtonHover = False And bMaximiseButtonClicked = False Then
                picForm.Line (UserControl.Width - 980, UserControl.Extender.Top + iVertical)-(UserControl.Width - 590, UserControl.Extender.Top + iVertical), lButtonGradientTop
            ElseIf bMaximiseButtonHover = True And bMaximiseButtonClicked = False Then
                picForm.Line (UserControl.Width - 980, UserControl.Extender.Top + iVertical)-(UserControl.Width - 590, UserControl.Extender.Top + iVertical), lButtonGradientTopHover
            ElseIf bMaximiseButtonHover = True And bMaximiseButtonClicked = True Then
                picForm.Line (UserControl.Width - 980, UserControl.Extender.Top + iVertical)-(UserControl.Width - 590, UserControl.Extender.Top + iVertical), lButtonGradientTopClicked
            End If
            iVertical = iVertical + 15
        Next

        ' Button Bottom Gradient Base Colour
        For i = 1 To 7
            If bMaximiseButtonHover = False And bMaximiseButtonClicked = False Then
                picForm.Line (UserControl.Width - 980, UserControl.Extender.Top + iVertical)-(UserControl.Width - 590, UserControl.Extender.Top + iVertical), lButtonGradientBottom(i)
            ElseIf bMaximiseButtonHover = True And bMaximiseButtonClicked = False Then
                picForm.Line (UserControl.Width - 980, UserControl.Extender.Top + iVertical)-(UserControl.Width - 590, UserControl.Extender.Top + iVertical), lButtonGradientBottomHover(i)
            ElseIf bMaximiseButtonHover = True And bMaximiseButtonClicked = True Then
                picForm.Line (UserControl.Width - 980, UserControl.Extender.Top + iVertical)-(UserControl.Width - 590, UserControl.Extender.Top + iVertical), lButtonGradientBottomClicked(i)
            End If
            iVertical = iVertical + 15
        Next
        
        ' Draw the Maximise Button Display
        ' Inside Button Display
        If myForm.WindowState <> 2 Then
            picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 150)-(UserControl.Width - 705, UserControl.Extender.Top + 150), &HFFFFFF
            picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 165)-(UserControl.Width - 705, UserControl.Extender.Top + 165), &HFFFFFF
            picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 180)-(UserControl.Width - 705, UserControl.Extender.Top + 180), &HFFFFFF
            picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 195)-(UserControl.Width - 705, UserControl.Extender.Top + 195), &HE9E9E9
            picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 210)-(UserControl.Width - 705, UserControl.Extender.Top + 210), &HE2E2E2
            picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 225)-(UserControl.Width - 705, UserControl.Extender.Top + 225), &HDCDCDC
            picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 240)-(UserControl.Width - 705, UserControl.Extender.Top + 240), &HD7D7D7
            
            ' Outside Borders
            If xVisualStyles = VistaLite Then
                picForm.Line (UserControl.Width - 870, UserControl.Extender.Top + 135)-(UserControl.Width - 690, UserControl.Extender.Top + 135), &H9E8A7C    ' Top Border
                picForm.Line (UserControl.Width - 870, UserControl.Extender.Top + 255)-(UserControl.Width - 690, UserControl.Extender.Top + 255), &H9E8A7C    ' Bottom Border
            End If
            
            picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 135)-(UserControl.Width - 705, UserControl.Extender.Top + 135), &H665653    ' Top Border
            picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 255)-(UserControl.Width - 705, UserControl.Extender.Top + 255), &H665653    ' Bottom Border
            picForm.Line (UserControl.Width - 870, UserControl.Extender.Top + 150)-(UserControl.Width - 870, UserControl.Extender.Top + 255), &H665653    ' Left Border
            picForm.Line (UserControl.Width - 705, UserControl.Extender.Top + 150)-(UserControl.Width - 705, UserControl.Extender.Top + 255), &H665653    ' Right Border
    
            picForm.Line (UserControl.Width - 825, UserControl.Extender.Top + 180)-(UserControl.Width - 735, UserControl.Extender.Top + 180), &H665653    ' Top Border
            picForm.Line (UserControl.Width - 825, UserControl.Extender.Top + 215)-(UserControl.Width - 735, UserControl.Extender.Top + 215), &H665653    ' Bottom Border
            picForm.Line (UserControl.Width - 825, UserControl.Extender.Top + 180)-(UserControl.Width - 825, UserControl.Extender.Top + 215), &H665653    ' Left Border
            picForm.Line (UserControl.Width - 750, UserControl.Extender.Top + 180)-(UserControl.Width - 750, UserControl.Extender.Top + 215), &H665653    ' Right Border
            
            If bMaximiseButtonHover = False And bMaximiseButtonClicked = False Then
                picForm.Line (UserControl.Width - 810, UserControl.Extender.Top + 195)-(UserControl.Width - 750, UserControl.Extender.Top + 195), lButtonGradientBottom(1)
            ElseIf bMaximiseButtonHover = True And bMaximiseButtonClicked = False Then
                picForm.Line (UserControl.Width - 810, UserControl.Extender.Top + 195)-(UserControl.Width - 750, UserControl.Extender.Top + 195), lButtonGradientBottomHover(1)
            ElseIf bMaximiseButtonHover = True And bMaximiseButtonClicked = True Then
                picForm.Line (UserControl.Width - 810, UserControl.Extender.Top + 195)-(UserControl.Width - 750, UserControl.Extender.Top + 195), lButtonGradientBottomClicked(1)
            End If
            
            ' Outside Borders
            If xVisualStyles = VistaLite Then
                picForm.Line (UserControl.Width - 870, UserControl.Extender.Top + 135)-(UserControl.Width - 690, UserControl.Extender.Top + 135), &H9E8A7C    ' Top Border
                picForm.Line (UserControl.Width - 870, UserControl.Extender.Top + 255)-(UserControl.Width - 690, UserControl.Extender.Top + 255), &H9E8A7C    ' Bottom Border
            End If
        Else
            picForm.Line (UserControl.Width - 810, UserControl.Extender.Top + 135)-(UserControl.Width - 705, UserControl.Extender.Top + 135), &HFFFFFF
            picForm.Line (UserControl.Width - 810, UserControl.Extender.Top + 150)-(UserControl.Width - 795, UserControl.Extender.Top + 150), &HFFFFFF
            picForm.Line (UserControl.Width - 795, UserControl.Extender.Top + 150)-(UserControl.Width - 720, UserControl.Extender.Top + 150), lButtonGradientBottom(1)
            picForm.Line (UserControl.Width - 720, UserControl.Extender.Top + 150)-(UserControl.Width - 705, UserControl.Extender.Top + 150), &HFFFFFF
            picForm.Line (UserControl.Width - 720, UserControl.Extender.Top + 165)-(UserControl.Width - 705, UserControl.Extender.Top + 165), &HFFFFFF
            picForm.Line (UserControl.Width - 720, UserControl.Extender.Top + 180)-(UserControl.Width - 705, UserControl.Extender.Top + 180), &HFFFFFF
            picForm.Line (UserControl.Width - 735, UserControl.Extender.Top + 195)-(UserControl.Width - 705, UserControl.Extender.Top + 195), &HFFFFFF
            picForm.Line (UserControl.Width - 735, UserControl.Extender.Top + 150)-(UserControl.Width - 735, UserControl.Extender.Top + 195), lButtonGradientBottom(1)
            
            picForm.Line (UserControl.Width - 810, UserControl.Extender.Top + 120)-(UserControl.Width - 705, UserControl.Extender.Top + 120), &H665653    ' Top Border
            picForm.Line (UserControl.Width - 735, UserControl.Extender.Top + 210)-(UserControl.Width - 705, UserControl.Extender.Top + 210), &H665653    ' Bottom Border
            picForm.Line (UserControl.Width - 825, UserControl.Extender.Top + 120)-(UserControl.Width - 825, UserControl.Extender.Top + 170), &H665653    ' Left Border
            picForm.Line (UserControl.Width - 705, UserControl.Extender.Top + 120)-(UserControl.Width - 705, UserControl.Extender.Top + 225), &H665653    ' Right Border

            picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 180)-(UserControl.Width - 750, UserControl.Extender.Top + 180), &HFFFFFF
            picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 195)-(UserControl.Width - 840, UserControl.Extender.Top + 195), &HFFFFFF
            picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 195)-(UserControl.Width - 840, UserControl.Extender.Top + 195), &HFFFFFF
            picForm.Line (UserControl.Width - 765, UserControl.Extender.Top + 195)-(UserControl.Width - 750, UserControl.Extender.Top + 195), &HFFFFFF
            picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 210)-(UserControl.Width - 840, UserControl.Extender.Top + 210), &HFFFFFF
            picForm.Line (UserControl.Width - 765, UserControl.Extender.Top + 210)-(UserControl.Width - 750, UserControl.Extender.Top + 210), &HFFFFFF
            picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 225)-(UserControl.Width - 840, UserControl.Extender.Top + 225), &HFFFFFF
            picForm.Line (UserControl.Width - 765, UserControl.Extender.Top + 225)-(UserControl.Width - 750, UserControl.Extender.Top + 225), &HFFFFFF
            picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 240)-(UserControl.Width - 750, UserControl.Extender.Top + 240), &HFFFFFF

            picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 165)-(UserControl.Width - 750, UserControl.Extender.Top + 165), &H665653    ' Top Border
            picForm.Line (UserControl.Width - 855, UserControl.Extender.Top + 255)-(UserControl.Width - 750, UserControl.Extender.Top + 255), &H665653    ' Bottom Border
            picForm.Line (UserControl.Width - 870, UserControl.Extender.Top + 165)-(UserControl.Width - 870, UserControl.Extender.Top + 270), &H665653    ' Left Border
            picForm.Line (UserControl.Width - 750, UserControl.Extender.Top + 165)-(UserControl.Width - 750, UserControl.Extender.Top + 270), &H665653    ' Right Border
        End If
    End If
Errhandler:
End Sub

Private Sub picForm_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call moveForm_MouseUp(Button, Shift, x, y)
End Sub

Private Sub picForm_Paint()
    ' Paints the Form header and label
    If bPaintForm = False Then
        ' doi sang chu unicode
        'sFormCaption = ToUni(sFormCaption)
        
        ' Set the Form header bottom colour
        Col = lFormGradientBottom
        lBottomR = (Col And &HFF&)
        lBottomG = (Col And &HFF00&) / &H100
        lBottomB = (Col And &HFF0000) / &H10000

        ' Set the Form header top colour
        Col = lFormGradientTop
        lTopR = (Col And &HFF&)
        lTopG = (Col And &HFF00&) / &H100
        lTopB = (Col And &HFF0000) / &H10000

        ' Clear the Form picturebox for drawing and apply the gradient colour
        picForm.Cls
        Set picForm.Picture = CreateGradient(UserControl.Width / Screen.TwipsPerPixelX, UserControl.Height / Screen.TwipsPerPixelY, False, RGB(lTopR, lTopG, lTopB), RGB(lBottomR, lBottomG, lBottomB), RGBBlend)
        
        ' Display the Text and Icon
        picForm.FontBold = bFontBold
        picForm.FontItalic = bFontItalic
        picForm.FontSize = dFontSize
        picForm.FontStrikethru = bFontStrikeThru
        picForm.FontUnderline = bFontUnderline
        If bDisplayIcon = False Then
            picForm.CurrentX = 90
        Else
            picForm.ScaleMode = 1
            If imgFormPic.Picture <> 0 Then
                picForm.CurrentX = 330
                picForm.PaintPicture imgFormPic.Picture, 60, 60, 240, 240
            Else
                picForm.CurrentX = 90
            End If
        End If
        picForm.CurrentY = (picForm.Height - picForm.TextHeight(sFormCaption)) / 2
    ' hien caption
        'picForm.Print
        TextOutWP picForm.hdc, 10, 8, StrPtr(sFormCaption), Len(sFormCaption)
        
        ' Draw the Buttons
        Call DrawMinimiseButton
        Call DrawMaximiseButton
        Call DrawCloseButton
        
        ' Top Border Line
        picForm.Line (0, 0)-(picForm.Width - 15, 0), lFormOuterBorder
        
        ' Left Border Line
        picForm.Line (0, 0)-(0, picForm.Height), lFormOuterBorder
        picForm.Line (15, 15)-(15, picForm.Height), lFormInnerBorder
        
        ' Right Border Line
        picForm.Line (picForm.Width - 30, 15)-(picForm.Width - 30, picForm.Height), lFormMiddleBorder
        picForm.Line (picForm.Width - 15, 0)-(picForm.Width - 15, picForm.Height), lFormOuterBorder
        
        If myForm.WindowState <> 2 Then
            ' Draw the Mask Colours
            ' Top Left Border Mask
            picForm.Line (0, 0)-(75, 0), &HFF00FF
            picForm.Line (0, 15)-(45, 15), &HFF00FF
            picForm.Line (0, 30)-(30, 30), &HFF00FF
            picForm.Line (0, 45)-(15, 45), &HFF00FF
            picForm.Line (0, 60)-(15, 60), &HFF00FF

            ' Top Left Border
            picForm.Line (45, 15)-(75, 15), lFormOuterBorder
            picForm.Line (30, 30)-(45, 30), lFormOuterBorder
            picForm.Line (15, 45)-(15, 75), lFormOuterBorder
            
            ' Top Right Border Mask
            picForm.Line (picForm.Width - 75, 0)-(picForm.Width, 0), &HFF00FF
            picForm.Line (picForm.Width - 45, 15)-(picForm.Width, 15), &HFF00FF
            picForm.Line (picForm.Width - 30, 30)-(picForm.Width, 30), &HFF00FF
            picForm.Line (picForm.Width - 15, 45)-(picForm.Width, 45), &HFF00FF
            picForm.Line (picForm.Width - 15, 60)-(picForm.Width, 60), &HFF00FF
    
            ' Top Right Border
            picForm.Line (picForm.Width - 75, 15)-(picForm.Width - 45, 15), lFormOuterBorder
            picForm.Line (picForm.Width - 45, 30)-(picForm.Width - 30, 30), lFormOuterBorder
            picForm.Line (picForm.Width - 30, 45)-(picForm.Width - 15, 75), lFormOuterBorder
    
            UserControl.Picture = picForm.Image
            myForm.Picture = UserControl.Picture
            Call MakeTransparent(myForm, &HFF00FF)
        End If
        
        Call UserControl_Paint
        bPaintForm = True
    End If
End Sub

Private Sub tmrMouseMove_Timer()
    Dim pt As POINTAPI

    ' See where the cursor is.
    GetCursorPos pt

    ' Translate into window coordinates.
    If WindowFromPointXY(pt.x, pt.y) <> picForm.hWnd Then
        bCloseButtonHover = False
        bCloseButtonClicked = False
        bMinimiseButtonHover = False
        bMaximiseButtonClicked = False
        bMaximiseButtonHover = False
        bMinimiseButtonClicked = False

        If bMouseOnForm = False Then
            ' Draw the Buttons
            Call DrawMinimiseButton
            Call DrawMaximiseButton
            Call DrawCloseButton
            bMouseOnForm = True
        End If
    Else
        bMouseOnForm = False

        ' Draw the Buttons
        Call DrawMinimiseButton
        Call DrawMaximiseButton
        Call DrawCloseButton
    End If
End Sub

Private Sub UserControl_Initialize()
    ' Initialise the default values
    bCloseButton = True
    bCloseButtonClicked = False
    bCloseButtonHover = False
    bDisplayIcon = False
    bFontBold = False
    bFontItalic = False
    bFontStrikeThru = False
    bFontUnderline = False
    bMaximiseButton = True
    bMaximiseButtonClicked = False
    bMaximiseButtonHover = False
    bMinimiseButton = True
    bMinimiseButtonClicked = False
    bMinimiseButtonHover = False
    bMouseClicked = False
    bMouseOnForm = False
    bPaintForm = False
    bRightClick = False
    bSystemTray = False
    bTransparency = False
    bUnloadForm = False
    dFontSize = 8
    iNumControls = 0
    iTransparency = 15
    lFormMaxHeight = 0
    lFormMinHeight = 0
    lFormMaxWidth = 0
    lFormMinWidth = 0
    lSysTrayMenu = 0
    
    ' TwipX and TwipY used only for easier writing
    TwipX = Screen.TwipsPerPixelX
    TwipY = Screen.TwipsPerPixelY
    
    ' Following variables used to speed-up the process (prevent recalc of BorderPixels * ...)
    BorderWidth = BorderPixels * TwipX
    BorderHeight = BorderPixels * TwipY
    
    xVisualStyles = VistaLite
End Sub

Private Sub UserControl_InitProperties()
    Call UserControlsCreate
    Set myForm = UserControl.Parent
    myForm.BorderStyle = 0
    myForm.AutoRedraw = True
    
    Call SelectColorScheme
    Call UserControl_Paint
    
    Set myForm.Icon = Nothing
    picForm.Width = picForm.Width - 1
    sFormCaption = UserControl.Extender.Name
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Call moveForm_MouseMove(Button, Shift, x, y)
    
    Dim Message As Long
    Message = x / Screen.TwipsPerPixelX
    
    Select Case Message
        Case WM_LBUTTONDBLCLK
            moveForm.Visible = Not moveForm.Visible
            moveForm.WindowState = Abs(Not moveForm.Visible)
        Case WM_RBUTTONUP
            Dim pt As POINTAPI
            Dim StartX As Single
            Dim StartY As Single
            bRightClick = True
            
            With pt  ' The x & y position where you want the menu displayed
                .x = x
                .y = y
                
                GetCursorPos pt
                RaiseEvent Execute(TrackPopupMenuEx(lSysTrayMenu, TPM_RETURNCMD, .x, .y, myForm.hWnd, ByVal 0&))
            End With
            bRightClick = False
            
            On Error GoTo Errhandler
            If bUnloadForm = True Then
                Call moveForm_Unload(0)
            End If
    End Select
    Exit Sub

Errhandler:
    
End Sub

Private Sub UserControl_Paint()
    On Error GoTo Errhandler
    Set myForm = UserControl.Parent
    myForm.BorderStyle = 0
    
    If bTransparency = True Then Call MakeSemiTransparent(UserControl.Parent.hWnd, iTransparency)
    
    ' Draw the Form Border Lines
    myForm.Cls
    
    ' Left Form Border Line
    myForm.Line (0, 0)-(0, myForm.Height), lFormOuterBorder
    myForm.Line (15, 0)-(15, myForm.Height), lFormInnerBorder
    myForm.Line (30, 0)-(30, myForm.Height), lFormGradientBottom
    myForm.Line (45, 0)-(45, myForm.Height), lFormGradientBottom
    
    ' Right Border Line
    myForm.Line (myForm.Width - 15, 0)-(myForm.Width - 15, myForm.Height), lFormOuterBorder
    myForm.Line (myForm.Width - 30, 0)-(myForm.Width - 30, myForm.Height), lFormMiddleBorder
    myForm.Line (myForm.Width - 45, 0)-(myForm.Width - 45, myForm.Height), lFormGradientBottom
    myForm.Line (myForm.Width - 60, 0)-(myForm.Width - 60, myForm.Height), lFormGradientBottom
    
    ' Bottom Border Line
    myForm.Line (30, myForm.Height - 60)-(myForm.Width - 30, myForm.Height - 60), lFormGradientBottom
    myForm.Line (30, myForm.Height - 45)-(myForm.Width - 30, myForm.Height - 45), lFormGradientBottom
    myForm.Line (15, myForm.Height - 30)-(myForm.Width - 15, myForm.Height - 30), lFormMiddleBorder
    myForm.Line (0, myForm.Height - 15)-(myForm.Width - 15, myForm.Height - 15), lFormOuterBorder
Errhandler:
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    Set moveForm = UserControl.Parent
    Set myForm = UserControl.Parent
    myForm.AutoRedraw = True
    
    ' Load saved properties
    xVisualStyles = PropBag.ReadProperty("Style", 0)
    Call SelectColorScheme
    Call UserControlsCreate
    sFormCaption = PropBag.ReadProperty("Caption", UserControl.Extender.Name)
    bDisplayIcon = PropBag.ReadProperty("DisplayIcon", False)
    picForm.Font = PropBag.ReadProperty("Font", Ambient.Font)
    bFontBold = PropBag.ReadProperty("FontBold", Ambient.Font)
    bFontItalic = PropBag.ReadProperty("FontItalic", Ambient.Font)
    dFontSize = PropBag.ReadProperty("FontSize", Ambient.Font)
    bFontStrikeThru = PropBag.ReadProperty("FontStrikeThru", Ambient.Font)
    bFontUnderline = PropBag.ReadProperty("FontUnderline", Ambient.Font)
    lFormCaptionColor = PropBag.ReadProperty("ForeColor", &H0&)
    imgFormPic.Picture = PropBag.ReadProperty("Icon", Nothing)
    lFormMaxHeight = PropBag.ReadProperty("MaxHeight", 0)
    lFormMaxWidth = PropBag.ReadProperty("MaxWidth", 0)
    lFormMinHeight = PropBag.ReadProperty("MinHeight", 0)
    lFormMinWidth = PropBag.ReadProperty("MinWidth", 0)
    bCloseButton = PropBag.ReadProperty("ShowCloseButton", True)
    bMinimiseButton = PropBag.ReadProperty("ShowMinimiseButton", True)
    bMaximiseButton = PropBag.ReadProperty("ShowMaximiseButton", True)
    bSystemTray = PropBag.ReadProperty("ShowSytemTrayIcon", False)
    xVisualStyles = PropBag.ReadProperty("Style", 0)
    bTransparency = PropBag.ReadProperty("Transparency", False)
    iTransparency = PropBag.ReadProperty("TransparencyLevel", 15)
    
    Call ShowInTheTaskbar(myForm.hWnd, True)
    If bTransparency = True Then Call MakeSemiTransparent(UserControl.Parent.hWnd, iTransparency)
    picForm.ForeColor = lFormCaptionColor
    Call UserControl_Paint
    Call UserControl_Resize
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    Call RemoveTransparent(myForm)
    
    If bUnloadForm = False Then
        ' Create the required controls
        Call UserControlsCreate
        
        UserControl.Height = 390
        UserControl.Extender.Align = 1
       
        ' Position the Form header
        picForm.Move 0, 0, UserControl.Width, UserControl.Height
        picForm.ZOrder 0
        
        bPaintForm = False
        Call picForm_Paint
        bPaintForm = True
    End If
End Sub

Private Sub UserControl_Show()
    myForm.Refresh
End Sub

Private Sub UserControl_Terminate()
    If bUnloadForm = False Then
        myForm.Cls
        Set myForm.Picture = Nothing
    Else
        Unload moveForm
        Unload myForm
        If bSystemTray = True Then
            Call DeleteIconFromTray
        End If
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    ' Save properties
    Call PropBag.WriteProperty("Caption", sFormCaption, UserControl.Extender.Name)
    Call PropBag.WriteProperty("DisplayIcon", bDisplayIcon, False)
    Call PropBag.WriteProperty("Font", picForm.Font, Ambient.Font)
    Call PropBag.WriteProperty("FontBold", picForm.FontBold, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", lFormCaptionColor, &H0&)
    Call PropBag.WriteProperty("FontItalic", picForm.FontItalic, Ambient.Font)
    Call PropBag.WriteProperty("FontSize", picForm.FontSize, Ambient.Font)
    Call PropBag.WriteProperty("FontStrikethru", picForm.FontStrikethru, Ambient.Font)
    Call PropBag.WriteProperty("FontUnderline", picForm.FontUnderline, Ambient.Font)
    Call PropBag.WriteProperty("Icon", imgFormPic.Picture, Nothing)
    Call PropBag.WriteProperty("MaxHeight", lFormMaxHeight, 0)
    Call PropBag.WriteProperty("MaxWidth", lFormMaxWidth, 0)
    Call PropBag.WriteProperty("MinHeight", lFormMinHeight, 0)
    Call PropBag.WriteProperty("MinWidth", lFormMinWidth, 0)
    Call PropBag.WriteProperty("ShowCloseButton", bCloseButton, True)
    Call PropBag.WriteProperty("ShowMinimiseButton", bMinimiseButton, True)
    Call PropBag.WriteProperty("ShowMaximiseButton", bMaximiseButton, True)
    Call PropBag.WriteProperty("ShowSytemTrayIcon", bSystemTray, False)
    Call PropBag.WriteProperty("Style", xVisualStyles, 0)
    Call PropBag.WriteProperty("Transparency", bTransparency, False)
    Call PropBag.WriteProperty("TransparencyLevel", iTransparency, 15)
End Sub

Public Sub AddSysTrayItem(ByVal MenuID As Long, ByVal MenuCaption As String, Optional bDefault As Boolean = False, Optional bChecked As Boolean = False, Optional eItemState As menuEStates)
    If lSysTrayMenu = 0 Then
        lSysTrayMenu = CreatePopupMenu()
    End If
    
    ' Add a SysTray menu item
    If MenuCaption <> "-" Then
        Call AppendMenu(lSysTrayMenu, MFS_STRING Or -bChecked * MFS_CHECKED, MenuID, ByVal MenuCaption)
    Else
        Call AppendMenu(lSysTrayMenu, MFS_STRING Or MFS_SEPARATOR, MenuID, ByVal vbNullString)
    End If
    
    ' Default item
    If bDefault Then Call SetMenuDefaultItem(lSysTrayMenu, MenuID, 0)
    
    ' Disabled (Regular color text)
    If eItemState = xDisabled Then Call EnableMenuItem(lSysTrayMenu, MenuID, MFS_BYCOMMAND Or MFS_DISABLED)
    ' Disabled (Disabled color text)
    If eItemState = xGrayed Then Call EnableMenuItem(lSysTrayMenu, MenuID, MFS_BYCOMMAND Or MFS_GRAYED)
End Sub

Public Sub AmendSysTrayItem(ByVal MenuID As Long, Optional bDefault As Boolean = False, Optional bChecked As Boolean = False, Optional eItemState As menuEStates)
    ' Amend a SysTray menu item
    Dim mnuItemInfo As MENUITEMINFO
    Dim hMenu As Long
    Dim hSubMenu As Long
    Dim BuffStr As String * 80   ' Define as largest possible menu text.
    
    With mnuItemInfo
        .cbSize = Len(mnuItemInfo)   ' 44
        .dwTypeData = BuffStr & Chr(0)
        .fType = MFS_STRING
        .cch = Len(mnuItemInfo.dwTypeData)   ' 80
        .fMask = MIIM_STATE
        .wID = MenuID
    End With
    
    ' Get the MenuID original details
    Call GetMenuItemInfo(lSysTrayMenu, MenuID, False, mnuItemInfo)
    
    ' Check or Uncheck
    If bChecked = False Then
        mnuItemInfo.fState = mnuItemInfo.fState - MFS_CHECKED
    Else
        mnuItemInfo.fState = (mnuItemInfo.fState Or MFS_CHECKED)
    End If
    
    ' Default or Not Default
    If bDefault = False Then
        mnuItemInfo.fState = mnuItemInfo.fState - MFS_DEFAULT
    Else
        mnuItemInfo.fState = (mnuItemInfo.fState Or MFS_DEFAULT)
    End If
    
    ' Set the MenuID new details
    Call SetMenuItemInfo(lSysTrayMenu, MenuID, False, mnuItemInfo)
    
    ' Disabled (Regular color text)
    If eItemState = xDisabled Then Call EnableMenuItem(lSysTrayMenu, MenuID, MFS_BYCOMMAND Or MFS_DISABLED)
    ' Disabled (Disabled color text)
    If eItemState = xGrayed Then Call EnableMenuItem(lSysTrayMenu, MenuID, MFS_BYCOMMAND Or MFS_GRAYED)
    ' Enabled if none of the above
    If eItemState <> xDisabled And eItemState <> xGrayed Then Call EnableMenuItem(lSysTrayMenu, MenuID, MFS_BYCOMMAND Or MFS_ENABLED)
End Sub

