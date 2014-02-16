VERSION 5.00
Begin VB.UserControl ListBox 
   ClientHeight    =   4590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3645
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   ScaleHeight     =   306
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   243
   Begin VB.VScrollBar Bar 
      Height          =   4095
      Left            =   2760
      TabIndex        =   0
      Top             =   360
      Width           =   255
   End
End
Attribute VB_Name = "ListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------
' AeroListBox ActiveX Control
' Based on ucCoolList 1.2 by Carles P.V. [CodeId=29586]
'--------------------------------------------------------------------------
' Copyright © 2007-2008 by Fauzie's Software. All rights reserved.
'--------------------------------------------------------------------------
' Author : Fauzie
' E-Mail : fauzie811@yahoo.com
'--------------------------------------------------------------------------

Option Explicit

'========================================================================================
' Subclasser declarations
'========================================================================================

Private Enum eMsgWhen
    [MSG_AFTER] = 1                                                           'Message calls back after the original (previous) WndProc
    [MSG_BEFORE] = 2                                                          'Message calls back before the original (previous) WndProc
    [MSG_BEFORE_AND_AFTER] = MSG_AFTER Or MSG_BEFORE                          'Message calls back before and after the original (previous) WndProc
End Enum

Private Type tSubData                                                         'Subclass data type
    hwnd                   As Long                                            'Handle of the window being subclassed
    nAddrSub               As Long                                            'The address of our new WndProc (allocated memory).
    nAddrOrig              As Long                                            'The address of the pre-existing WndProc
    nMsgCntA               As Long                                            'Msg after table entry count
    nMsgCntB               As Long                                            'Msg before table entry count
    aMsgTblA()             As Long                                            'Msg after table array
    aMsgTblB()             As Long                                            'Msg Before table array
End Type

Private sc_aSubData()      As tSubData                                        'Subclass data array
Private Const ALL_MESSAGES As Long = -1                                       'All messages added or deleted
Private Const GMEM_FIXED   As Long = 0                                        'Fixed memory GlobalAlloc flag
Private Const GWL_WNDPROC  As Long = -4                                       'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04     As Long = 88                                       'Table B (before) address patch offset
Private Const PATCH_05     As Long = 93                                       'Table B (before) entry count patch offset
Private Const PATCH_08     As Long = 132                                      'Table A (after) address patch offset
Private Const PATCH_09     As Long = 137                                      'Table A (after) entry count patch offset

Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

'//

Public Enum AlignmentCts
    [AlignLeft]
    [AlignCenter]
    [AlignRight]
End Enum

Public Enum AppearanceCts
    [Flat]
    [3D]
End Enum

Public Enum BorderStyleCts
    [None]
    [Fixed Single]
End Enum

Public Enum OrderTypeCts
    [Ascendent]
    [Descendent]
End Enum

Public Enum SelectModeCts
    [Single]
    [Multiple]
End Enum

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT2) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT2, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT2, ByVal hBrush As Long) As Long
Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hDC As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function DrawTextW Lib "user32.dll" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, ByRef lpRect As RECT2, ByVal wFormat As Long) As Long ' Modified
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Type TRIVERTEX
    X          As Long
    Y          As Long
    R          As Integer
    G          As Integer
    B          As Integer
    Alpha      As Integer
End Type

Private Type RGB
    R          As Integer
    G          As Integer
    B          As Integer
End Type

Private Type GRADIENT_RECT
    UpperLeft  As Long
    LowerRight As Long
End Type

Private Type RECT2
    X1         As Long
    Y1         As Long
    X2         As Long
    Y2         As Long
End Type

Private Type POINTAPI
    X          As Long
    Y          As Long
End Type

Private Const PS_SOLID As Long = 0

Private Const GRADIENT_FILL_RECT_H As Long = &H0
Private Const GRADIENT_FILL_RECT_V As Long = &H1
Private Const DT_LEFT As Long = &H0
Private Const DT_CENTER As Long = &H1
Private Const DT_RIGHT As Long = &H2
Private Const DT_VCENTER As Long = &H4
Private Const DT_WORDBREAK As Long = &H10
Private Const DT_SINGLELINE As Long = &H20
Private Const WM_MOUSEWHEEL As Long = &H20A&

'-------------------------------------------------------------------------------------------
' UserControl constants / types / variables / events
'-------------------------------------------------------------------------------------------

Private Type tItem
    Text       As String
    Icon       As Integer
    IconSelected As Integer
End Type

Private m_List() As tItem                             ' List array of items (Text, icons)
Private m_Selected() As Boolean                       ' List array of items (Selected/Unselected)
Private m_nItems As Integer                           ' Number of Items

Private m_LastBar As Integer                          ' Last scroll bar value
Private m_LastItem As Integer                         ' Last Selected item
Private m_LastY As Single                             ' Last Y value [pixels] (prevents item repaint)
Private m_AnchorItemState As Boolean                  ' Anchor item value (multiple selection).
'  Case extended selection: all selected items
'  will be set to Anchor selection state.

Private m_EnsureVisible As Boolean                    ' Ensure visible last m_Selected item (ListIndex)

Private m_ItemRct() As RECT2                          ' Item rectangle
Private m_TextRct() As RECT2                          ' Item text rectangle
Private m_IconPt() As POINTAPI                        ' Item icon position

Private m_tmpItemHeight As Integer                    ' Item height [pixels]
Private m_VisibleRows As Integer                      ' Visible rows in control area
Private m_Scrolling As Boolean                        ' Scrolling by mouse
Private m_ScrollingY As Long                          ' Y Scrolling coordinate flag (scroll speed = f(Y))
Private m_HasFocus As Boolean                         ' Control has focus
Private m_Resizing As Boolean                         ' Prevent repaints when Resizing

Private m_pImgList As Object                          ' Will point to ImageList control
Private m_ILScale As Integer                          ' ImageList parent scale mode

Private m_ColorBack As Long                           ' Back color [Normal]
Private m_ColorBackSel As Long                        ' Back color [Selected]
Private m_ColorFont As Long                           ' Font color [Normal]
Private m_ColorFontSel As Long                        ' Font color [Selected]
Private m_ColorGradient1 As RGB                       ' Gradient color from [Selected]
Private m_ColorGradient2 As RGB                       ' Gradient color  to  [Selected]
Private m_ColorBox As Long                            ' Box border color

Private WithEvents m_Font As StdFont                  ' Font object
Attribute m_Font.VB_VarHelpID = -1

Private m_Alignment As AlignmentCts
Private m_Apeareance As AppearanceCts
Private m_BorderStyle As BorderStyleCts
Private m_BackNormal As OLE_COLOR
Private m_BackSelected As OLE_COLOR
Private m_Focus As Boolean
Private m_FontNormal As OLE_COLOR
Private m_FontSelected As OLE_COLOR
Private m_HoverSelection As Boolean
Private m_ItemHeight As Integer
Private m_ItemHeightAuto As Boolean
Private m_ItemOffset As Integer
Private m_ItemTextLeft As Integer
Private m_ListIndex As Integer
Private m_OrderType As OrderTypeCts
Private m_SelectMode As SelectModeCts
Private m_TopIndex As Integer

Private Const m_def_Appearance = 1
Private Const m_def_Alignment = DT_LEFT
Private Const m_def_BackNormal = vbWhite
Private Const m_def_BackSelected = &HFDF4E2
Private Const m_def_BorderStyle = 1
Private Const m_def_Focus = -1
Private Const m_def_FontNormal = vbBlack
Private Const m_def_FontSelected = vbBlack
Private Const m_def_HoverSelection = 0
Private Const m_def_ItemHeightAuto = -1
Private Const m_def_ItemOffset = 0
Private Const m_def_ItemTextLeft = 2
Private Const m_def_OrderType = 0
Private Const m_def_SelectMode = 0
Private Const m_def_WordWrap = -1

Public Event Click()
Attribute Click.VB_UserMemId = -600
Public Event DblClick()
Attribute DblClick.VB_UserMemId = -601
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_UserMemId = -602
Public Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_UserMemId = -603
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_UserMemId = -604
Public Event ListIndexChange()
Attribute ListIndexChange.VB_MemberFlags = "200"
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_UserMemId = -605
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_UserMemId = -606
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_UserMemId = -607
Public Event Scroll()
Public Event TopIndexChange()

'-------------------------------------------------------------------------------------------
' Init/Read/Write properties
'-------------------------------------------------------------------------------------------

Private Sub UserControl_InitProperties()

    UserControl.Appearance = m_def_Appearance
    m_BorderStyle = m_def_BorderStyle

    Set UserControl.Font = Ambient.Font
    Set m_Font = Ambient.Font

    m_FontNormal = m_def_FontNormal
    m_FontSelected = m_def_FontSelected
    m_BackNormal = m_def_BackNormal
    m_BackSelected = m_def_BackSelected

    m_Alignment = m_def_Alignment
    m_Focus = m_def_Focus
    m_HoverSelection = m_def_HoverSelection

    m_ItemHeight = UserControl.TextHeight("TextHeight")
    m_ItemHeightAuto = m_def_ItemHeightAuto
    m_ItemOffset = m_def_ItemOffset
    m_ItemTextLeft = m_def_ItemTextLeft

    m_OrderType = m_def_OrderType
    m_SelectMode = m_def_SelectMode

    m_ListIndex = -1
    m_TopIndex = -1

    SetColors

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", -1)

    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)

    m_FontNormal = PropBag.ReadProperty("FontNormal", m_def_FontNormal)
    m_FontSelected = PropBag.ReadProperty("FontSelected", m_def_FontSelected)
    m_BackNormal = PropBag.ReadProperty("BackNormal", m_def_BackNormal)
    UserControl.BackColor = PropBag.ReadProperty("BackNormal", m_def_BackNormal)
    m_BackSelected = PropBag.ReadProperty("BackSelected", m_def_BackSelected)

    m_Alignment = PropBag.ReadProperty("Alignment", m_def_Alignment)
    m_Focus = PropBag.ReadProperty("Focus", m_def_Focus)
    m_HoverSelection = PropBag.ReadProperty("HoverSelection", m_def_HoverSelection)

    m_ItemOffset = PropBag.ReadProperty("ItemOffset", m_def_ItemOffset)
    m_ItemHeightAuto = PropBag.ReadProperty("ItemHeightAuto", m_def_ItemHeightAuto)
    m_ItemTextLeft = PropBag.ReadProperty("ItemTextLeft", m_def_ItemTextLeft)

    m_OrderType = PropBag.ReadProperty("OrderType", m_def_OrderType)
    m_SelectMode = PropBag.ReadProperty("SelectMode", m_def_SelectMode)

    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Set UserControl.MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)

    Dim sTmp   As String
    sTmp = PropBag.ReadProperty("ItemHeight", 0)
    If (sTmp < UserControl.TextHeight("")) Then
        m_ItemHeight = UserControl.TextHeight("")
    Else                                              'NOT (STMP...
        m_ItemHeight = sTmp
    End If

    m_ListIndex = -1
    m_TopIndex = -1

    SetColors
    
    Call UserControl_Resize
    
    If Ambient.UserMode Then
      Call Subclass_Start(UserControl.hwnd)
      Call Subclass_AddMsg(hwnd, WM_MOUSEWHEEL)
    End If
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 1)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, 1)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, -1)

    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("FontNormal", m_FontNormal, m_def_FontNormal)
    Call PropBag.WriteProperty("FontSelected", m_FontSelected, m_def_FontSelected)
    Call PropBag.WriteProperty("BackNormal", m_BackNormal, m_def_BackNormal)
    Call PropBag.WriteProperty("BackSelected", m_BackSelected, m_def_BackSelected)

    Call PropBag.WriteProperty("Alignment", m_Alignment, m_def_Alignment)
    Call PropBag.WriteProperty("Focus", m_Focus, m_def_Focus)
    Call PropBag.WriteProperty("HoverSelection", m_HoverSelection, m_def_HoverSelection)

    Call PropBag.WriteProperty("ItemHeight", m_ItemHeight, 0)
    Call PropBag.WriteProperty("ItemHeightAuto", m_ItemHeightAuto, m_def_ItemHeightAuto)
    Call PropBag.WriteProperty("ItemOffset", m_ItemOffset, m_def_ItemOffset)
    Call PropBag.WriteProperty("ItemTextLeft", m_ItemTextLeft, m_def_ItemTextLeft)

    Call PropBag.WriteProperty("OrderType", m_OrderType, m_def_OrderType)
    Call PropBag.WriteProperty("SelectMode", m_SelectMode, m_def_SelectMode)

    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("MouseIcon", UserControl.MouseIcon, Nothing)

End Sub

'-------------------------------------------------------------------------------------------
'-- UserControl initialitation, focus, size, refresh, termination
'-------------------------------------------------------------------------------------------

Private Sub UserControl_Initialize()

'-- Initialize arrays

    ReDim m_List(0)
    ReDim m_Selected(0)
    '-- Initialize position flags
    m_EnsureVisible = -1                              ' Ensure visible last selected
    m_LastItem = -1                                   ' Last selected
    m_LastY = -1                                      ' Last Y coordinate
    '-- Initialize font object
    Set m_Font = New StdFont

End Sub

Private Sub UserControl_EnterFocus()

    m_HasFocus = -1
    DrawFocus m_ListIndex

End Sub

Private Sub UserControl_ExitFocus()

    m_HasFocus = 0
    DrawItem m_ListIndex

End Sub

Private Sub UserControl_Resize()
    On Error Resume Next

    '-- Set item height

    Dim Tmp As Integer
    Tmp = Abs(m_BorderStyle)
    
    If (m_ItemHeightAuto) Then
        m_tmpItemHeight = UserControl.TextHeight("")
    Else                                              '(M_ITEMHEIGHTAUTO) = FALSE/0
        If (m_ItemHeight < UserControl.TextHeight("")) Then
            m_tmpItemHeight = UserControl.TextHeight("")
        Else                                          'NOT (M_ITEMHEIGHT...
            m_tmpItemHeight = m_ItemHeight
        End If
    End If

    '-- Get visible rows and readjust control height
    m_VisibleRows = ScaleHeight \ m_tmpItemHeight
    Height = ((m_VisibleRows) * m_tmpItemHeight + (Tmp * 4)) * Screen.TwipsPerPixelX + (Height - ScaleHeight * Screen.TwipsPerPixelY)

    '-- Locate and resize drawing area, calc. rects and readjust scroll bar
    m_Resizing = -1
    With Bar
        .Move ScaleWidth - .Width - Tmp, Tmp, .Width, ScaleHeight - (Tmp * 2)
        .Visible = 0
    End With                                          'BAR
    ReDim m_ItemRct(m_VisibleRows - 1)
    ReDim m_TextRct(m_VisibleRows - 1)
    ReDim m_IconPt(m_VisibleRows - 1)
    CalculateRects
    ReadjustBar
    
    SetWindowRgn hwnd, CreateRoundRectRgn(0, 0, ScaleWidth + 1, ScaleHeight + 1, 2, 2), True
    
    m_Resizing = 0

    On Error GoTo 0
End Sub

Private Sub UserControl_Paint()

    If (Not Ambient.UserMode) Then

        UserControl.Cls

        Select Case m_Alignment
            Case 0
                UserControl.CurrentX = m_ItemTextLeft + m_ItemOffset + (m_BorderStyle * 2)
            Case 1
                UserControl.CurrentX = ((ScaleWidth - (m_BorderStyle * 4)) - UserControl.TextWidth(Ambient.DisplayName)) * 0.5
            Case 2
                UserControl.CurrentX = ((ScaleWidth - (m_BorderStyle * 4)) - UserControl.TextWidth(Ambient.DisplayName)) - m_ItemOffset
        End Select
        UserControl.CurrentY = m_ItemOffset + m_BorderStyle * 2

        SetTextColor UserControl.hDC, m_ColorFont
        UserControl.Print Ambient.DisplayName

    Else                                              'NOT (NOT...
        If (Not m_Resizing) Then DrawList
    End If

    If m_BorderStyle Then
        Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), RGB(165, 172, 178), B
    End If

End Sub

Private Sub UserControl_Terminate()
    
    On Error Resume Next
    Erase m_List
    Erase m_Selected
    Set m_pImgList = Nothing
    m_Scrolling = 0
    Call Subclass_StopAll

End Sub

'-------------------------------------------------------------------------------------------
'-- ScrollBar
'-------------------------------------------------------------------------------------------

Private Sub Bar_Change()

    If (m_LastBar <> Bar.Value) Then
        m_LastBar = Bar.Value
        m_LastY = -1
        If (m_ListIndex = m_LastItem) Then
            DrawList
        End If
        RaiseEvent Scroll
        RaiseEvent TopIndexChange
    End If

End Sub

Private Sub Bar_Scroll()

    Bar_Change
    RaiseEvent Scroll

End Sub

'-------------------------------------------------------------------------------------------
' Scrolling / Events
'-------------------------------------------------------------------------------------------

'-- Click()

Private Sub UserControl_Click()

    If (m_ListIndex > -1) Then RaiseEvent Click

End Sub

'-- DblClick()

Private Sub UserControl_DblClick()

    If (m_ListIndex > -1) Then RaiseEvent DblClick

End Sub

'-- KeyDown(KeyCode, Shift)

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    Debug.Print KeyCode
    If (m_nItems = 0 Or m_ListIndex = -1) Then
        RaiseEvent KeyDown(KeyCode, Shift)
        Exit Sub                                      '---> Bottom
    End If
    
    Select Case KeyCode
        Case 13                                       '{Enter}
            If (m_ListIndex > 0) Then ListIndex = ListIndex
            RaiseEvent Click

        Case 38                                       '{Up arrow}
            If (m_ListIndex > 0) Then ListIndex = ListIndex - 1

        Case 40                                       '{Down arrow}
            If (m_ListIndex < m_nItems - 1) Then ListIndex = ListIndex + 1

        Case 33                                       '{PageUp}
            If (m_ListIndex > m_VisibleRows) Then
                ListIndex = ListIndex - (m_VisibleRows - 1)
            Else                                      'NOT (M_LISTINDEX...
                ListIndex = 0
            End If

        Case 34                                       '{PageDown}
            If (m_ListIndex < m_nItems - m_VisibleRows - 1) Then
                ListIndex = ListIndex + (m_VisibleRows - 1)
            Else                                      'NOT (M_LISTINDEX...
                ListIndex = m_nItems - 1
            End If

        Case 36                                       '{Start}
            ListIndex = 0

        Case 35                                       '{End}
            ListIndex = m_nItems - 1

        Case 32                                       '{Space} Select/Unselect
            If (m_SelectMode <> 0 And m_ListIndex > -1) Then
                m_Selected(m_ListIndex) = Not m_Selected(m_ListIndex)
                DrawItem m_ListIndex
                DrawFocus m_ListIndex
            End If
            RaiseEvent Click
    End Select

    RaiseEvent KeyDown(KeyCode, Shift)

End Sub

'-- KeyPress(KeyAscii)

Private Sub UserControl_KeyPress(KeyAscii As Integer)

    RaiseEvent KeyPress(KeyAscii)

End Sub

'-- KeyPress(KeyCode, Shift)

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)

    RaiseEvent KeyUp(KeyCode, Shift)

End Sub

'-- MouseDown(Button, Shift, x, y)

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If (Button = vbRightButton) Then

        RaiseEvent MouseDown(Button, Shift, X, Y)
        Exit Sub                                      '---> Bottom
    End If

    Dim SelectedListIndex As Integer
    SelectedListIndex = Bar.Value + Int(Y / m_tmpItemHeight)

    If (SelectedListIndex >= 0 And SelectedListIndex < m_nItems) Then
        Select Case m_SelectMode
            Case 0                                    ' [Single]
                m_Selected(SelectedListIndex) = -1
            Case 1                                    ' [Multiple]
                m_Selected(SelectedListIndex) = Not m_Selected(SelectedListIndex)
                m_AnchorItemState = m_Selected(SelectedListIndex)
        End Select

        m_LastY = Y
        ListIndex = SelectedListIndex
    End If

    m_Scrolling = -1
    RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

'-- MouseMove(Button, Shift, x, y)

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim SelectedListIndex As Integer

    m_ScrollingY = Y

    If (Y < 0) Then
        ScrollUp
        RaiseEvent MouseMove(Button, Shift, X, Y)
        Exit Sub                                      '---> Bottom
    End If
    If (Y > ScaleHeight) Then
        ScrollDown
        RaiseEvent MouseMove(Button, Shift, X, Y)
        Exit Sub                                      '---> Bottom
    End If

    If (m_HoverSelection Or Button = 1) And (Y \ m_tmpItemHeight <> m_LastY \ m_tmpItemHeight) Then

        SelectedListIndex = Bar + (Y \ m_tmpItemHeight)

        If (SelectedListIndex >= 0 And SelectedListIndex < m_nItems) Then
            m_Selected(SelectedListIndex) = m_AnchorItemState
            ListIndex = SelectedListIndex
            m_LastY = Y
        End If
    End If

    RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

'-- MouseUp(Button, Shift, x, y)

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    m_Scrolling = 0
    m_AnchorItemState = -1
    RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

'========================================================================================
' UserControl subclass procedure
'========================================================================================

Public Sub zSubclass_Proc(ByVal bBefore As Boolean, _
                          ByRef bHandled As Boolean, _
                          ByRef lReturn As Long, _
                          ByRef lhWnd As Long, _
                          ByRef uMsg As Long, _
                          ByRef wParam As Long, _
                          ByRef lParam As Long _
                          )
Attribute zSubclass_Proc.VB_MemberFlags = "40"
                          
    Select Case lhWnd
    Case UserControl.hwnd
      Select Case uMsg
      Case WM_MOUSEWHEEL
        If Bar.Visible Then
          If (wParam = &H780000) Then
            If Bar.Value > 0 Then Bar.Value = Bar.Value - 1
          ElseIf (wParam = &HFF880000) Then
            If Bar.Value < Bar.Max Then Bar.Value = Bar.Value + 1
          End If
        End If
      End Select
    End Select
End Sub

'-------------------------------------------------------------------------------------------
' Methods
'-------------------------------------------------------------------------------------------



'-- SetImageList

Public Sub SetImageList(ImageListControl)

    Set m_pImgList = ImageListControl

    On Error Resume Next
    m_ILScale = m_pImgList.Parent.ScaleMode
    On Error GoTo 0

    UserControl_Paint

End Sub

'-- AddItem
'-- 0 , ... , n-1 [n = ListCount]

Public Sub AddItem(ByVal Text As Variant, _
                   Optional ByVal Icon As Integer = -1, _
                   Optional ByVal IconSelected As Integer = -1)

    m_List(m_nItems).Text = CStr(Text)
    m_List(m_nItems).Icon = Icon
    m_List(m_nItems).IconSelected = IconSelected
    If IconSelected = -1 Then m_List(m_nItems).IconSelected = Icon
    m_nItems = m_nItems + 1

    ReDim Preserve m_List(m_nItems)
    ReDim Preserve m_Selected(m_nItems)

    ReadjustBar
    If (m_nItems < m_VisibleRows + 1) Then
        DrawItem (m_nItems - 1)
    End If

End Sub

'-- InsertItem

Public Sub InsertItem(ByVal Index As Integer, _
                      ByVal Text As Variant, _
                      Optional ByVal Icon As Integer, _
                      Optional ByVal IconSelected As Integer)

    If (m_nItems = 0 Or Index > m_nItems) Then Err.Raise 381

    m_nItems = m_nItems + 1
    ReDim Preserve m_List(m_nItems)
    ReDim Preserve m_Selected(m_nItems)

    Dim i      As Long
    For i = m_nItems - 1 To Index Step -1
        m_List(i + 1) = m_List(i)
        m_Selected(i + 1) = m_Selected(i)
    Next i

    m_List(Index).Text = CStr(Text)
    m_List(Index).Icon = Icon
    m_List(Index).IconSelected = IconSelected
    m_Selected(Index) = 0

    ReadjustBar
    m_EnsureVisible = 0
    If (m_ListIndex > -1 And Index <= m_ListIndex) Then
        ListIndex = ListIndex + 1
    End If
    UserControl_Paint

End Sub

'-- ModifyItem

Public Sub ModifyItem(ByVal Index As Integer, _
                      Optional ByVal Text As Variant = vbEmpty, _
                      Optional ByVal Icon As Integer = -1, _
                      Optional ByVal IconSelected As Integer = -1)

    If (m_nItems = 0 Or Index > m_nItems) Then Err.Raise 381

    If (Text <> vbEmpty) Then m_List(Index).Text = CStr(Text)
    If (Icon > -1) Then m_List(Index).Icon = Icon
    If (IconSelected > -1) Then m_List(Index).IconSelected = IconSelected

    DrawItem Index
    DrawFocus m_ListIndex

End Sub

'-- RemoveItem

Public Sub RemoveItem(ByVal Index As Integer)

    If (m_nItems = 0 Or Index > m_nItems - 1) Then Err.Raise 381

    If (Index < m_nItems) Then
        Dim i  As Long
        For i = Index To m_nItems - 1
            m_List(i) = m_List(i + 1)
            m_Selected(i) = m_Selected(i + 1)
        Next i
    End If

    m_nItems = m_nItems - 1
    ReDim Preserve m_List(m_nItems)
    ReDim Preserve m_Selected(m_nItems)

    ReadjustBar
    m_EnsureVisible = 0

    If (Index < m_ListIndex) Then
        If (m_ListIndex > -1) Then ListIndex = ListIndex - 1
    ElseIf (Index = m_ListIndex) Then                 'NOT (INDEX...
        ListIndex = -1
    End If

    If (m_nItems < m_VisibleRows) Then
        UserControl.Cls
    End If
    UserControl_Paint

End Sub

'-- FindFirst

Public Function FindFirst(ByVal FindString As String, _
                          Optional ByVal StartIndex As Integer = 0, _
                          Optional ByVal StartWith As Boolean = 0) As Integer

    If (m_nItems = 0) Then Err.Raise 381

    Dim i      As Long
    For i = StartIndex To m_nItems
        If (StartWith) Then
            If (InStr(1, LCase$(m_List(i).Text), LCase$(FindString)) = 1) Then FindFirst = i: Exit Function
        Else                                          '(STARTWITH) = FALSE/0
            If (InStr(1, LCase$(m_List(i).Text), LCase$(FindString)) > 1) Then FindFirst = i: Exit Function
        End If
    Next i

    '-- FindString not found
    FindFirst = -1

End Function

'-- Clear

Public Sub Clear()

'-- Hide scroll bar

    Bar.Visible = 0
    Bar.Max = 0
    '-- Clear and resize drawing area
    UserControl.Cls
    'UserControl.Move 0, 0, ScaleWidth, ScaleHeight
    '-- Reset Item arrays
    ReDim m_List(0)
    ReDim m_Selected(0)
    m_nItems = 0

    m_LastItem = -1
    m_ListIndex = -1
    m_TopIndex = -1

End Sub

'-- Order

Public Sub Order()

    Dim i0     As Long
    Dim i1     As Long
    Dim i2     As Long
    Dim d      As Long
    Dim xItem  As tItem
    Dim bDesc  As Boolean

    If (m_nItems > 1) Then

        i0 = 0
        bDesc = (m_OrderType = [Descendent])

        If (m_SelectMode = [Single]) Then
            If (m_ListIndex > -1) Then m_Selected(m_ListIndex) = 0
        End If

        Do
            d = d * 3 + 1
        Loop Until d > m_nItems

        Do
            d = d \ 3
            For i1 = d + i0 To m_nItems + i0 - 1

                xItem = m_List(i1)
                i2 = i1

                Do While (m_List(i2 - d).Text > xItem.Text) Xor bDesc
                    m_List(i2) = m_List(i2 - d)
                    i2 = i2 - d
                    If (i2 - d < i0) Then Exit Do
                Loop
                m_List(i2) = xItem
            Next i1
        Loop Until d = 1

        ListIndex = -1
        Bar = 0

        '-- Unselect all and refresh
        ReDim m_Selected(0 To m_nItems)
        UserControl_Paint
    End If

End Sub

'-------------------------------------------------------------------------------------------
'-- Draw List / Item / Focus
'-------------------------------------------------------------------------------------------

'-- DrawList

Private Sub DrawList()

    Dim i      As Long

    If (Extender.Visible And UBound(m_List)) Then
        '-- Draw visible rows
        For i = Bar.Value To Bar.Value + m_VisibleRows - 1
            DrawItem i
        Next i
        '-- Draw focus
        DrawFocus m_ListIndex
    End If

End Sub

'-- DrawItem

Private Sub DrawItem(ByVal Index As Integer)

    Dim nRctIndex As Integer, mTmpRct As RECT2
    '-- Item out of area?

    If (Index < Bar.Value Or Index > Bar.Value + m_VisibleRows - 1) Then Exit Sub
    If (Index > UBound(m_List) - 1) Then Exit Sub

    nRctIndex = Index - Bar.Value

    '-- Draw m_Selected Item
    If (m_Selected(Index)) Then
        
        '-- Draw back area
        DrawBack UserControl.hDC, m_ItemRct(nRctIndex), m_ColorBackSel
        SetTextColor UserControl.hDC, m_ColorFontSel

        '-- Draw icon
        If (Not m_pImgList Is Nothing) Then
            On Error Resume Next                      'Image list icon # out of bounds
            m_pImgList.ListImages(m_List(Index).IconSelected).Draw UserControl.hDC, ScaleX(m_ItemOffset + 2 + (m_BorderStyle * 2), vbPixels, m_ILScale), ScaleY(m_ItemRct(nRctIndex).Y1 + (m_tmpItemHeight - m_pImgList.ImageHeight) * 0.5, vbPixels, m_ILScale), 1
            On Error GoTo 0
        End If
    Else                                              '(M_SELECTED(INDEX)) = FALSE/0

        '-- Draw back area
        DrawBack UserControl.hDC, m_ItemRct(nRctIndex), m_ColorBack
        SetTextColor UserControl.hDC, m_ColorFont

        '-- Draw icon
        If (Not m_pImgList Is Nothing) Then
            On Error Resume Next                      'Image list icon # out of bounds
            m_pImgList.ListImages(m_List(Index).Icon).Draw UserControl.hDC, ScaleX(m_ItemOffset + 2 + (m_BorderStyle * 2), vbPixels, m_ILScale), ScaleY(m_ItemRct(nRctIndex).Y1 + (m_tmpItemHeight - m_pImgList.ImageHeight) * 0.5, vbPixels, m_ILScale), 1
            On Error GoTo 0
        End If
    End If
    
    LSet mTmpRct = m_TextRct(nRctIndex)
    On Error Resume Next
    If (Not m_pImgList Is Nothing) Then mTmpRct.X1 = mTmpRct.X1 + m_pImgList.ImageWidth + 2 + (m_BorderStyle * 2)
    
    '-- Draw text...
    DrawTextW UserControl.hDC, StrPtr(m_List(Index).Text), Len(m_List(Index).Text), mTmpRct, DT_SINGLELINE Or DT_VCENTER

End Sub

'-- DrawFocus

Private Sub DrawFocus(Index As Integer)

    If Not Me.Focus Then Exit Sub                     ' (Not m_Focus Or Not m_HasFocus) Then Exit Sub

'-- Item out of area ?
    If (Index < Bar.Value Or Index > Bar.Value + m_VisibleRows - 1) Then Exit Sub

    '-- Draw it
    UserControl.ForeColor = ShiftColorOXP(m_ColorBackSel, -700)
    RoundRect UserControl.hDC, m_ItemRct(Index - Bar.Value).X1, m_ItemRct(Index - Bar.Value).Y1, m_ItemRct(Index - Bar.Value).X2, m_ItemRct(Index - Bar.Value).Y2, 2, 2

End Sub

Private Sub DrawBack(ByVal hDC As Long, pRect As RECT2, ByVal Color As Long)

    Dim hBrush As Long

    hBrush = CreateSolidBrush(Color)
    FillRect hDC, pRect, hBrush
    DeleteObject hBrush

End Sub

Private Sub DrawBackGrad(ByVal hDC As Long, pRect As RECT2, Color1 As RGB, Color2 As RGB, ByVal Direction As Long)

    Dim v(1)   As TRIVERTEX
    Dim GRct   As GRADIENT_RECT

    '-- from

    With v(0)
        .X = pRect.X1
        .Y = pRect.Y1
        .R = Color1.R
        .G = Color1.G
        .B = Color1.B
        .Alpha = 0
    End With                                          'V(0)
    '-- to
    With v(1)
        .X = pRect.X2
        .Y = pRect.Y2
        .R = Color2.R
        .G = Color2.G
        .B = Color2.B
        .Alpha = 0
    End With                                          'V(1)

    GRct.UpperLeft = 0
    GRct.LowerRight = 1

    GradientFillRect hDC, v(0), 2, GRct, 1, Direction

End Sub

Private Sub ReadjustBar()
    On Error Resume Next

    If (m_nItems > m_VisibleRows) Then

        If (Not Bar.Visible) Then
            '-- Show scroll bar
            Bar.Visible = -1
            Bar.Refresh
            Bar.LargeChange = m_VisibleRows
            '-- Update item rects. right margin
            RigthOffsetRects Bar.Width
            '-- Repaint control area
            UserControl_Paint
        End If

    Else                                              'NOT (M_NITEMS...
        '-- Hide scroll bar
        Bar.Visible = 0
        '-- Update item rects. right margin
        RigthOffsetRects 0
    End If

    '-- Update Bar max value
    Bar.Max = m_nItems - m_VisibleRows

    On Error GoTo 0

End Sub

Private Sub CalculateRects()

    Dim i      As Long, tBS As Integer
    tBS = m_BorderStyle

    For i = 0 To m_VisibleRows - 1
        SetRect m_ItemRct(i), tBS * 2, i * m_tmpItemHeight + (tBS * 2), UserControl.ScaleWidth, i * m_tmpItemHeight + m_tmpItemHeight + (tBS * 2)
        SetRect m_TextRct(i), m_ItemOffset + m_ItemTextLeft + (tBS * 2), i * m_tmpItemHeight + m_ItemOffset + (tBS * 2), UserControl.ScaleWidth - m_ItemOffset, i * m_tmpItemHeight + m_tmpItemHeight - m_ItemOffset + (tBS * 2)
        m_IconPt(i).X = m_ItemOffset
        m_IconPt(i).Y = m_ItemOffset
    Next i

End Sub

Private Sub RigthOffsetRects(ByVal Offset As Long)

    Dim i      As Long, tBS As Integer
    tBS = m_BorderStyle

    For i = 0 To m_VisibleRows - 1
        m_ItemRct(i).X2 = UserControl.ScaleWidth - Offset - (tBS * 2)
        m_TextRct(i).X2 = UserControl.ScaleWidth - m_ItemOffset - Offset - (tBS * 2)
    Next i

End Sub

'-------------------------------------------------------------------------------------------
' Scroll Up/Down by mouse / multiple select
'-------------------------------------------------------------------------------------------

'-- ScrollUp

Private Sub ScrollUp()

    Dim t      As Long                                ' Timer counter
    Dim d      As Long                                ' Scrolling delay

    d = 500 + 20 * m_ScrollingY
    If (d < 40) Then d = 40

    '-- Scroll while MouseDown and mouse pos. < "Control top"
    Do While m_Scrolling And m_ScrollingY < 0
        If (GetTickCount - t > d) Then
            t = GetTickCount
            If (m_ListIndex > 0) Then
                If (m_SelectMode = [Multiple]) Then
                    m_Selected(m_ListIndex - 1) = m_AnchorItemState
                End If
                ListIndex = ListIndex - 1
            End If
        End If
        DoEvents
    Loop

End Sub

'-- ScrollDown

Private Sub ScrollDown()

    Dim t      As Long                                ' Timer counter
    Dim d      As Long                                ' Scrolling delay

    d = 500 - 20 * (m_ScrollingY - ScaleHeight - 1)
    If (d < 40) Then d = 40

    '-- Scroll while MouseDown and mouse pos. > "Control bottom"
    Do While m_Scrolling And m_ScrollingY > ScaleHeight - 1
        If (GetTickCount - t > d) Then
            t = GetTickCount
            If (m_ListIndex < m_nItems - 1) Then
                If (m_SelectMode = [Multiple]) Then
                    m_Selected(m_ListIndex + 1) = m_AnchorItemState
                End If
                ListIndex = ListIndex + 1
            End If
        End If
        DoEvents
    Loop

End Sub

'-------------------------------------------------------------------------------------------
' Colors
'-------------------------------------------------------------------------------------------

'-- SetColors

Private Sub SetColors()

    m_ColorBack = GetLngColor(m_BackNormal)
    m_ColorBackSel = GetLngColor(m_BackSelected)
    m_ColorFont = GetLngColor(m_FontNormal)
    m_ColorFontSel = GetLngColor(m_FontSelected)

End Sub

Private Function GetLngColor(Color As Long) As Long

    If (Color And &H80000000) Then
        GetLngColor = GetSysColor(Color And &H7FFFFFFF)
    Else                                              'NOT (COLOR...
        GetLngColor = Color
    End If

End Function

Private Function GetRGBColors(Color As Long) As RGB

    Dim HexColor As String

    HexColor = String$(6 - Len(Hex$(Color)), "0") & Hex$(Color)
    GetRGBColors.R = "&H" & Mid$(HexColor, 5, 2) & "00"
    GetRGBColors.G = "&H" & Mid$(HexColor, 3, 2) & "00"
    GetRGBColors.B = "&H" & Mid$(HexColor, 1, 2) & "00"

End Function

Private Function GetRGB(Color As Long) As RGB
'   Returns the RGB color value of the specified color.
    GetRGB.R = Color And 255
    GetRGB.G = (Color \ 256) And 255
    GetRGB.B = (Color \ 65536) And 255
    
End Function

Private Function ShiftColor(Color As Long, PercentInDecimal As Single) As Long
'   Add or remove a certain color quantity by how many percent.
    Dim RGB1 As RGB
        RGB1 = GetRGB(Color)
        
        RGB1.R = RGB1.R + PercentInDecimal * 255 ' Percent should already
        RGB1.G = RGB1.G + PercentInDecimal * 255 ' be translated.
        RGB1.B = RGB1.B + PercentInDecimal * 255 ' Ex. 50% -> 50 / 100 = 0.5
        
    If (PercentInDecimal > 0) Then ' RGB values must be between 0-255 only
        If (RGB1.R > 255) Then RGB1.R = 255
        If (RGB1.G > 255) Then RGB1.G = 255
        If (RGB1.B > 255) Then RGB1.B = 255
    Else
        If (RGB1.R < 0) Then RGB1.R = 0
        If (RGB1.G < 0) Then RGB1.G = 0
        If (RGB1.B < 0) Then RGB1.B = 0
    End If
    
    ShiftColor = RGB(RGB1.R, RGB1.G, RGB1.B) ' Return shifted color value
    
End Function

Private Function ShiftColorOXP(ByVal theColor As Long, Optional ByVal Base As Long = &HB0) As Long
  Dim Red As Long, Blue As Long, Green As Long
  Dim Delta As Long
  
  Blue = ((theColor \ &H10000) Mod &H100)
  Green = ((theColor \ &H100) Mod &H100)
  Red = (theColor And &HFF)
  Delta = &HFF - Base
  
  Blue = Base + Blue * Delta \ &HFF
  Green = Base + Green * Delta \ &HFF
  Red = Base + Red * Delta \ &HFF
  
  If Red > 255 Then Red = 255
  If Green > 255 Then Green = 255
  If Blue > 255 Then Blue = 255
  
  ShiftColorOXP = Red + 256& * Green + 65536 * Blue
End Function

'-------------------------------------------------------------------------------------------
' Properties
'-------------------------------------------------------------------------------------------

'-- Alignment

Public Property Get Alignment() As AlignmentCts
Attribute Alignment.VB_ProcData.VB_Invoke_Property = ";Appearance"

    Alignment = m_Alignment

End Property

Public Property Let Alignment(ByVal New_Alignment As AlignmentCts)

    m_Alignment = New_Alignment
    UserControl_Paint

End Property

'-- Appearance

Public Property Get Appearance() As AppearanceCts
Attribute Appearance.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Appearance.VB_UserMemId = -520

    Appearance = UserControl.Appearance

End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceCts)

    UserControl.Appearance() = New_Appearance

End Property

'-- BackNormal

Public Property Get BackNormal() As OLE_COLOR
Attribute BackNormal.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BackNormal.VB_UserMemId = -501

    BackNormal = m_BackNormal

End Property

Public Property Let BackNormal(ByVal New_BackNormal As OLE_COLOR)

    m_BackNormal = New_BackNormal
    m_ColorBack = GetLngColor(m_BackNormal)
    UserControl.BackColor = m_ColorBack
    UserControl_Paint

End Property

'-- BackSelected

Public Property Get BackSelected() As OLE_COLOR
Attribute BackSelected.VB_ProcData.VB_Invoke_Property = ";Appearance"

    BackSelected = m_BackSelected

End Property

Public Property Let BackSelected(ByVal New_BackSelected As OLE_COLOR)

    m_BackSelected = New_BackSelected
    m_ColorBackSel = GetLngColor(m_BackSelected)
    UserControl_Paint

End Property

'-- BorderStyle

Public Property Get BorderStyle() As BorderStyleCts
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BorderStyle.VB_UserMemId = -503

    BorderStyle = m_BorderStyle

End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleCts)

    m_BorderStyle = New_BorderStyle
    Call UserControl_Resize

End Property

'-- Enabled

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Enabled.VB_UserMemId = -514

    Enabled = UserControl.Enabled

End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)

    UserControl.Enabled() = New_Enabled
    Bar.Enabled = New_Enabled

End Property

'-- Focus

Public Property Get Focus() As Boolean
Attribute Focus.VB_ProcData.VB_Invoke_Property = ";Misc"

    Focus = m_Focus

End Property

Public Property Let Focus(ByVal New_Focus As Boolean)

    m_Focus = New_Focus
    If (New_Focus) Then
        DrawFocus m_ListIndex
    Else                                              '(NEW_FOCUS) = FALSE/0
        DrawItem m_ListIndex
    End If

End Property

'-- Font

Public Property Get Font() As Font
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512

    Set Font = m_Font

End Property

Public Property Set Font(ByVal New_Font As Font)

    With m_Font
        .Name = New_Font.Name
        .Size = New_Font.Size
        .Bold = New_Font.Bold
        .Italic = New_Font.Italic
        .Underline = New_Font.Underline
        .Strikethrough = New_Font.Strikethrough
    End With                                          'M_FONT
    UserControl_Paint

End Property

Private Sub m_Font_FontChanged(ByVal PropertyName As String)

    Set UserControl.Font = m_Font
    UserControl_Resize

End Sub

'-- FontNormal

Public Property Get FontNormal() As OLE_COLOR
Attribute FontNormal.VB_ProcData.VB_Invoke_Property = ";Font"

    FontNormal = m_FontNormal

End Property

Public Property Let FontNormal(ByVal New_FontNormal As OLE_COLOR)

    m_FontNormal = New_FontNormal
    m_ColorFont = GetLngColor(m_FontNormal)
    SetTextColor UserControl.hDC, m_ColorFont
    UserControl_Paint

End Property

'-- FontSelected

Public Property Get FontSelected() As OLE_COLOR
Attribute FontSelected.VB_ProcData.VB_Invoke_Property = ";Font"

    FontSelected = m_FontSelected

End Property

Public Property Let FontSelected(ByVal New_FontSelected As OLE_COLOR)

    m_FontSelected = New_FontSelected
    m_ColorFontSel = GetLngColor(m_FontSelected)
    UserControl_Paint

End Property

'-- HoverSelection

Public Property Get HoverSelection() As Boolean
Attribute HoverSelection.VB_ProcData.VB_Invoke_Property = ";Misc"

    HoverSelection = m_HoverSelection

End Property

Public Property Let HoverSelection(ByVal New_HoverSelection As Boolean)

    m_HoverSelection = New_HoverSelection
    DrawItem m_ListIndex
    DrawFocus m_ListIndex

End Property

'-- ItemHeight

Public Property Get ItemHeight() As Integer
Attribute ItemHeight.VB_ProcData.VB_Invoke_Property = ";Misc"

    ItemHeight = m_ItemHeight

End Property

Public Property Let ItemHeight(ByVal New_ItemHeight As Integer)

    m_ItemHeight = New_ItemHeight
    UserControl_Resize
    UserControl_Paint

End Property

'-- ItemHeightAuto

Public Property Get ItemHeightAuto() As Boolean

    ItemHeightAuto = m_ItemHeightAuto

End Property

Public Property Let ItemHeightAuto(ByVal New_ItemHeightAuto As Boolean)

    m_ItemHeightAuto = New_ItemHeightAuto
    UserControl_Resize
    UserControl_Paint

End Property

'-- ItemOffset

Public Property Get ItemOffset() As Integer

    ItemOffset = m_ItemOffset

End Property

Public Property Let ItemOffset(ByVal New_ItemOffset As Integer)

    If (New_ItemOffset <= m_tmpItemHeight) Then
        m_ItemOffset = New_ItemOffset
    End If
    CalculateRects
    If (Bar.Visible) Then RigthOffsetRects Bar.Width
    UserControl_Paint

End Property

'-- ItemTextLeft

Public Property Get ItemTextLeft() As Integer

    ItemTextLeft = m_ItemTextLeft

End Property

Public Property Let ItemTextLeft(ByVal New_ItemTextLeft As Integer)

    m_ItemTextLeft = New_ItemTextLeft
    CalculateRects
    If (Bar.Visible) Then RigthOffsetRects Bar.Width
    UserControl_Paint

End Property

'-- <ListCount>

Public Property Get ListCount() As Integer
Attribute ListCount.VB_ProcData.VB_Invoke_Property = ";List"
Attribute ListCount.VB_MemberFlags = "400"

    ListCount = m_nItems

End Property

'-- ListIndex

Public Property Get ListIndex() As Integer
Attribute ListIndex.VB_ProcData.VB_Invoke_Property = ";List"
Attribute ListIndex.VB_MemberFlags = "400"

    ListIndex = m_ListIndex

End Property

Public Property Let ListIndex(ByVal New_ListIndex As Integer)

    If (New_ListIndex < -1 Or New_ListIndex > m_nItems - 1) Then Err.Raise 380

    If (New_ListIndex < 0 Or m_nItems = 0) Then
        m_ListIndex = -1
        m_LastY = -1
    Else                                              'NOT (NEW_LISTINDEX...
        m_ListIndex = New_ListIndex
    End If

    '-- Unselect last / Select actual [Single selection mode]
    If (m_SelectMode = [Single]) Then
        If (m_LastItem > -1) Then m_Selected(m_LastItem) = 0
        If (m_ListIndex > -1) Then m_Selected(m_ListIndex) = -1
    End If

    '-- Draw last (delete Focus) ...
    DrawItem m_LastItem
    m_LastItem = m_ListIndex
    '-- ... and draw actual (draw Focus)
    DrawItem m_ListIndex
    DrawFocus m_ListIndex

    '-- Ensure visible actual Selected item
    If (m_EnsureVisible) Then
        If (m_ListIndex < Bar.Value And m_ListIndex > -1) Then
            Bar = m_ListIndex
        ElseIf (m_ListIndex > Bar.Value + m_VisibleRows - 1) Then    'NOT (M_LISTINDEX...
            Bar = m_ListIndex - m_VisibleRows + 1
        End If
    Else                                              '(M_ENSUREVISIBLE) = FALSE/0
        m_EnsureVisible = -1
    End If

    RaiseEvent ListIndexChange

End Property

'-- MouseIcon

Public Property Get MouseIcon() As Picture

    Set MouseIcon = UserControl.MouseIcon

End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)

    Set UserControl.MouseIcon = New_MouseIcon

End Property

'-- MousePointer

Public Property Get MousePointer() As MousePointerConstants

    MousePointer = UserControl.MousePointer

End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)

    UserControl.MousePointer() = New_MousePointer

End Property

'-- OrderType

Public Property Get OrderType() As OrderTypeCts

    OrderType = m_OrderType

End Property

Public Property Let OrderType(ByVal New_OrderType As OrderTypeCts)

    m_OrderType = New_OrderType

End Property

'-- <SelectedCount>

Public Property Get SelectedCount() As Integer
Attribute SelectedCount.VB_MemberFlags = "400"

    Dim i      As Long

    SelectedCount = 0
    For i = 0 To m_nItems
        If (m_Selected(i)) Then SelectedCount = SelectedCount + 1
    Next i

End Property

'-- SelectMode

Public Property Get SelectMode() As SelectModeCts

    SelectMode = m_SelectMode

End Property

Public Property Let SelectMode(ByVal New_SelectMode As SelectModeCts)

    Dim i      As Long
    m_SelectMode = New_SelectMode

    If (Ambient.UserMode) Then
        If (New_SelectMode = [Single]) Then
            '-- Unselect all and select actual
            If (m_ListIndex > -1) Then
                For i = LBound(m_List) To m_nItems
                    If (i <> m_ListIndex) Then m_Selected(i) = 0
                Next i
                m_Selected(m_ListIndex) = -1
                DrawItem m_ListIndex
                DrawFocus m_ListIndex
            End If
        End If
    End If

    ReadjustBar
    UserControl_Paint

End Property

'-- TopIndex

Public Property Get TopIndex() As Integer
Attribute TopIndex.VB_MemberFlags = "400"

    TopIndex = Bar

End Property

Public Property Let TopIndex(ByVal New_TopIndex As Integer)

    If (New_TopIndex < 0 Or New_TopIndex > m_nItems - m_VisibleRows) Then Err.Raise 380

    m_TopIndex = New_TopIndex
    Bar = New_TopIndex

    RaiseEvent TopIndexChange

End Property

'Last revised: 02/07/02
'-------------------------------------------------------------------------------------------
' Some methods passed to R/W properties:
'
' GetItem i    GetIcon i    GetIconSelected i    IsSelected i
' to           to           to                   to
' ItemText(i)  ItemIcon(i)  ItemIconSelected(i)  ItemSelected(i)
'
' Or use ModifyItem to change all item parameters at time

'-- ItemText

Public Property Get ItemText(ByVal Index As Integer) As String

    If (m_nItems = 0 Or Index > m_nItems) Then Err.Raise 381
    ItemText = m_List(Index).Text

End Property

Public Property Let ItemText(ByVal Index As Integer, ByVal Data As String)

    If (m_nItems = 0 Or Index > m_nItems) Then Err.Raise 381
    m_List(Index).Text = CStr(Data)
    DrawItem Index
    DrawFocus m_ListIndex

End Property

'-- ItemIcon

Public Property Get ItemIcon(ByVal Index As Integer) As Integer

    If (m_nItems = 0 Or Index > m_nItems) Then Err.Raise 381
    ItemIcon = m_List(Index).Icon

End Property

Public Property Let ItemIcon(ByVal Index As Integer, ByVal Data As Integer)

    If (m_nItems = 0 Or Index > m_nItems) Then Err.Raise 381
    m_List(Index).Icon = Data
    DrawItem Index
    DrawFocus m_ListIndex

End Property

'-- ItemIconSelected

Public Property Get ItemIconSelected(ByVal Index As Integer) As Integer

    If (m_nItems = 0 Or Index > m_nItems) Then Err.Raise 381
    ItemIconSelected = m_List(Index).IconSelected

End Property

Public Property Let ItemIconSelected(ByVal Index As Integer, ByVal Data As Integer)

    If (m_nItems = 0 Or Index > m_nItems) Then Err.Raise 381
    m_List(Index).IconSelected = Data
    DrawItem Index
    DrawFocus m_ListIndex

End Property

'-- ItemSelected

Public Property Get ItemSelected(ByVal Index As Integer) As Boolean

    If (m_nItems = 0 Or Index > m_nItems) Then Err.Raise 381
    ItemSelected = m_Selected(Index)

End Property

Public Property Let ItemSelected(ByVal Index As Integer, ByVal Data As Boolean)

    If (m_nItems = 0 Or Index > m_nItems) Then Err.Raise 381

    Select Case Data
        Case -1
            If (m_SelectMode = [Single]) Then
                ListIndex = Index
            Else                                      'NOT (M_SELECTMODE...
                m_Selected(Index) = -1
                DrawItem Index
                If (Index = m_ListIndex) Then DrawFocus Index
            End If
        Case 0
            If (m_SelectMode = [Single]) Then
            Else                                      'NOT (M_SELECTMODE...
                m_Selected(Index) = 0
                DrawItem Index
                If (Index = m_ListIndex) Then DrawFocus Index
            End If
    End Select

End Property

Public Sub SimulateKeyPress(KeyCode As Integer, Shift As Integer)

    UserControl_KeyDown KeyCode, Shift

End Sub

'========================================================================================
'Subclass routines below here - The programmer may call any of the following Subclass_??? routines
'========================================================================================

Private Sub Subclass_AddMsg(ByVal lhWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
  
    With sc_aSubData(zIdx(lhWnd))
        If (When And eMsgWhen.MSG_BEFORE) Then
            Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
        End If
        If (When And eMsgWhen.MSG_AFTER) Then
            Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
        End If
    End With
End Sub

Private Function Subclass_InIDE() As Boolean
    Debug.Assert zSetTrue(Subclass_InIDE)
End Function

Private Function Subclass_Start(ByVal lhWnd As Long) As Long

  Const CODE_LEN              As Long = 202
  Const FUNC_CWP              As String = "CallWindowProcA"
  Const FUNC_EBM              As String = "EbMode"
  Const FUNC_SWL              As String = "SetWindowLongA"
  Const MOD_USER              As String = "user32"
  Const MOD_VBA5              As String = "vba5"
  Const MOD_VBA6              As String = "vba6"
  Const PATCH_01              As Long = 18
  Const PATCH_02              As Long = 68
  Const PATCH_03              As Long = 78
  Const PATCH_06              As Long = 116
  Const PATCH_07              As Long = 121
  Const PATCH_0A              As Long = 186
  Static aBuf(1 To CODE_LEN)  As Byte
  Static pCWP                 As Long
  Static pEbMode              As Long
  Static pSWL                 As Long
  Dim i                       As Long
  Dim J                       As Long
  Dim nSubIdx                 As Long
  Dim sHex                    As String
  
    If (aBuf(1) = 0) Then
  
        sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"
        i = 1
        Do While J < CODE_LEN
            J = J + 1
            aBuf(J) = Val("&H" & Mid$(sHex, i, 2))
            i = i + 2
        Loop
    
        If (Subclass_InIDE) Then
            aBuf(16) = &H90
            aBuf(17) = &H90
            pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)
            If (pEbMode = 0) Then
                pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)
            End If
        End If
    
        pCWP = zAddrFunc(MOD_USER, FUNC_CWP)
        pSWL = zAddrFunc(MOD_USER, FUNC_SWL)
        ReDim sc_aSubData(0 To 0) As tSubData
      Else
        nSubIdx = zIdx(lhWnd, True)
        If (nSubIdx = -1) Then
            nSubIdx = UBound(sc_aSubData()) + 1
            ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData
        End If
    
        Subclass_Start = nSubIdx
    End If

    With sc_aSubData(nSubIdx)
        .hwnd = lhWnd
        .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)
        .nAddrOrig = SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrSub)
        Call RtlMoveMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)
        Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)
        Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)
        Call zPatchRel(.nAddrSub, PATCH_03, pSWL)
        Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)
        Call zPatchRel(.nAddrSub, PATCH_07, pCWP)
        Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))
    End With
End Function

Private Sub Subclass_Stop(ByVal lhWnd As Long)
  
    With sc_aSubData(zIdx(lhWnd))
        Call SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrOrig)
        Call zPatchVal(.nAddrSub, PATCH_05, 0)
        Call zPatchVal(.nAddrSub, PATCH_09, 0)
        Call GlobalFree(.nAddrSub)
        .hwnd = 0
        .nMsgCntB = 0
        .nMsgCntA = 0
        Erase .aMsgTblB()
        Erase .aMsgTblA()
    End With
End Sub

Private Sub Subclass_StopAll()
  
  Dim i As Long
  
    i = UBound(sc_aSubData())
    Do While i >= 0
        With sc_aSubData(i)
            If (.hwnd <> 0) Then
                Call Subclass_Stop(.hwnd)
            End If
        End With
        i = i - 1
    Loop
End Sub

'----------------------------------------------------------------------------------------
'These z??? routines are exclusively called by the Subclass_??? routines.
'----------------------------------------------------------------------------------------

Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
  
  Dim nEntry  As Long
  Dim nOff1   As Long
  Dim nOff2   As Long
  
    If (uMsg = ALL_MESSAGES) Then
        nMsgCnt = ALL_MESSAGES
      Else
        Do While nEntry < nMsgCnt
            nEntry = nEntry + 1
            If (aMsgTbl(nEntry) = 0) Then
                aMsgTbl(nEntry) = uMsg
                Exit Sub
            ElseIf (aMsgTbl(nEntry) = uMsg) Then
                Exit Sub
            End If
        Loop

        nMsgCnt = nMsgCnt + 1
        ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long
        aMsgTbl(nMsgCnt) = uMsg
    End If

    If (When = eMsgWhen.MSG_BEFORE) Then
        nOff1 = PATCH_04
        nOff2 = PATCH_05
      Else
        nOff1 = PATCH_08
        nOff2 = PATCH_09
    End If

    If (uMsg <> ALL_MESSAGES) Then
        Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))
    End If
    Call zPatchVal(nAddr, nOff2, nMsgCnt)
End Sub

Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
    zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
    Debug.Assert zAddrFunc
End Function

Private Function zIdx(ByVal lhWnd As Long, Optional ByVal bAdd As Boolean = False) As Long

    zIdx = UBound(sc_aSubData)
    Do While zIdx >= 0
        With sc_aSubData(zIdx)
            If (.hwnd = lhWnd) Then
                If (Not bAdd) Then
                    Exit Function
                End If
            ElseIf (.hwnd = 0) Then
                If (bAdd) Then
                    Exit Function
                End If
            End If
        End With
        zIdx = zIdx - 1
    Loop
  
    If (Not bAdd) Then
        Debug.Assert False
    End If
End Function

Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
    Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
    Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
    zSetTrue = True
    bValue = True
End Function

