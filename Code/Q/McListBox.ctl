VERSION 5.00
Begin VB.UserControl McListBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2205
   FillColor       =   &H80000018&
   FillStyle       =   0  'Solid
   ScaleHeight     =   247
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   147
End
Attribute VB_Name = "McListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^Gtech^Creations^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^¶¶^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^¶¶¶^^^^¶¶¶^^^^^^^^¶¶^^^^^¶¶^^^^^^^^^¶¶^^^¶¶¶¶¶¶^^^^^^^^^^^^^^^^^^^^^^^^¶¶^^^^^^^^¶¶¶¶¶^^^^^$
'$^^^^^¶¶¶^^^^¶¶¶^^^^^^^^¶¶^^^^^^^^^^^^^^^^¶¶^^^¶¶^^^¶¶^^^^^^^^^^^^^^^^^^^^^¶¶¶¶^^^^^^^¶^^^^¶¶^^^^$
'$^^^^^¶^¶¶^^¶^¶¶^^¶¶¶¶^^¶¶^^^^^¶¶^^¶¶¶¶^^¶¶¶¶¶^¶¶^^^¶¶^^¶¶¶¶¶^^¶¶^^¶¶^^^^^^^^¶¶^^^^^^^¶^^^^¶¶^^^^$
'$^^^^^¶^¶¶^^¶^¶¶^¶¶^^^¶^¶¶^^^^^¶¶^¶¶^^^¶^^¶¶^^^¶¶^^^¶¶^¶¶^^^¶¶^¶¶^^¶¶^^^^^^^^¶¶^^^^^^^^^^^^¶¶^^^^$
'$^^^^^¶^^¶¶¶^^¶¶^¶¶^^^^^¶¶^^^^^¶¶^¶¶¶^^^^^¶¶^^^¶¶¶¶¶¶^^¶¶^^^¶¶^^¶¶¶¶^^^^^^^^^¶¶^^^^^^^^^^^¶¶^^^^^$
'$^^^^^¶^^¶¶¶^^¶¶^¶¶^^^^^¶¶^^^^^¶¶^^¶¶¶¶^^^¶¶^^^¶¶^^^¶¶^¶¶^^^¶¶^^^¶¶^^^^^^^^^^¶¶^^^^^^^^^^¶¶^^^^^^$
'$^^^^^¶^^^¶^^^¶¶^¶¶^^^^^¶¶^^^^^¶¶^^^^¶¶¶^^¶¶^^^¶¶^^^¶¶^¶¶^^^¶¶^^¶¶¶¶^^^^^^^^^¶¶^^^^^^^^^¶¶^^^^^^^$
'$^^^^^¶^^^¶^^^¶¶^¶¶^^^¶^¶¶^^^^^¶¶^¶^^^¶¶^^¶¶^^^¶¶^^^¶¶^¶¶^^^¶¶^¶¶^^¶¶^^^^^^^^¶¶^^^^¶¶^^¶¶^^^^^^^^$
'$^^^^^¶^^^^^^^¶¶^^¶¶¶¶^^¶¶¶¶¶¶^¶¶^^¶¶¶¶^^^^¶¶¶^¶¶¶¶¶¶^^^¶¶¶¶¶^^¶¶^^¶¶^^^^^^¶¶¶¶¶¶^^¶¶^¶¶¶¶¶¶¶^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^By^Jim^Jose^^^^^^^Email^jimjosev33@yahoo.com^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

'-----------------------------------------------------------------------------------------------------------
' SourceCode : McListBox 3.2
' Auther     : Jim Jose
' Email      : jimjosev33@yahoo.com
' Date       : 3-9-2005
' Purpose    : An upgraded version of VBListBox with Icons, Item HighLight and many more
' CopyRight  : JimJose © Gtech Creations - 2005
'-----------------------------------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------------------------------
' About :
'        This control is a replacement for the vb's inbuilt 'ListBox'
' control. This version comes with,

'   1. Unicode' support
'   2. MultiSelect Option
'   3. Custom Icons for each item
'   4. Item highlight by making the fond bold
'   5. Nice gradient effects
'   6. Custom RowHeight
'   7. SortItems(Ascendindg/Descending)
'   8. Grid Lines
'   9. API SCrollBars
'  10. XP style
'  11. Mouse Wheel
'  12. And many more...
'
'       This is my second attempt on this purpose. The first one
' 'ListBoxEx' is there on PSC. The primary objective of this update
' was multiple icons and Item HighLight. I tried to implement most
' of the suggessions I got on previous submission.
'
'       But, this version also will not support 'Multiple Columns' and
' 'Column headers'. I agree.. that will not cause much effort. But, this
' control is ment for easy use in small & nice projects. So I like to
' keep it small!!
'
'       This control also uses 'McImageList', by which the images can store
' externally and make referance to them. This way we can greatly reduse the
' memory use caused by a big 'Picture Array', where each item needs a seperate
' picture object.
'
' Special Note:
'       Please 'Refresh' the listbox after 'item adding'. I didn't include
' 'Sorting' and 'Redrawing' on 'ItemAdd' for better speed..

'-----------------------------------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------------------------------
'
' Credits/Thanks :
'
'    Paul Carton    -   For his unbeatable self subclasser.
'    Gary Noble     -   For his API ScrollBars and ColorBlending code.
'    Carls P.V      -   For his excellent DIB-gradient routine.
'    Dana Seaman    -   The Master of Unicode support.
'
'-----------------------------------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------------------------------
' Updates :
'
'   4/9/2005 - New version Submitted to PSC!!
'   4/9/2005 - Deleted from PSC due to some bug found in the DEMO project.
'   5/9/2004 - Bug fixed!!, Resubmitted...
'   6/9/2005 - Fixed an error found in 'ListBold' and 'ListIcon' properties
'
' Version 1.6 :
'   7/9/2005 - A new developed version with 'Extended' and 'Simple' multiselection
'              option. 'Shift' key for extended selection and 'Ctrl' for simple selection
'
' Version 1.7 :
'   10/9/2005 - New version released with Gary Nobles's (Phantom Man) color blending code for
'              XP style selection colors. There is a great improvement in the
'              look and feel of the control. Also fixed some minor bugs!
'              Added functions  - SelectAll, ClearSelection, Find Selection ( all by Gary )
'
'              This version uses Paul Carton's Selft 'SubClassing' code, for getting the
'              mousewheel functionality
'
'              Also included 'Gary Nobles' excellent API scrollbars, which made the control
'              more powerful than ever!!. Give full credit to Gary, who continuesly helped me
'              improving the functionality and fixing errors. Thanks Gary!!
'
' Version 1.8 :
'   14/9/2005 - Multiselection bug,noticed by Matt is fixed! (see the comment section of page
'               for details). Also fixed the color identifying bug when selecting *System Color*
'               as informed by Richard.
'
' Version 1.9 :
'   15/9/2005 - Added the functionality for "FlatScrollBars" and back "Picture".
'               Also included a property "AutoHideScrollBars", which desides the
'               hiding of scrollbars when item count is less the the number of
'               items that can show at a time (Scalewidth/rowHeight).
'
'               Also fixed the scrollbar enabling + subclassing issue when control
'               creates dynamically.
' Version 2.2
'   21/9/2005 - Improved version with ItemCompleter. When mouse moves over an item
'               which have a bigger size(width) that can't be showed fully on the list,
'               the Item completer window will be shown with the full text
'
' Version 2.3
'   28/9/2005 - Added functionality to show "MultiLine" "ItemCompleter"
'
' Version 3.2
'   10-2-2006 - Released the new hybrid version 3.2 which is the combination of
'                   1) ListBox
'                   2) DriveList
'                   3) FolderList
'                   4) FileList
'                   5) File Browser
'                   6) Folder Browser
'               You can select the mode to any of these....
'
'               NB : Apart from the vb's inbuilt controls McListBox (hybrid) can
'                    show files with it's ICON. The icons are extracted directly from the
'                    file such that we get the same icon that explorer showing...
'                    * You can select the size for extracted icon tooo...

'              IMPORTANT:
'                   * In the new version of McListBox, it is not neccessary that you refresh the
'                   control to diaply the list after adding the items. The control will automatically
'                   Refresh after an item has been added.
'                   * It is done by an internal timer and thus avoids repeated refreshing
'                   when items are adding from a loop (still 5 times faster than vbListBox)
'
' Version 3.3
'   12-2-2006 - Added features... 1) Horizontal scrollbars
'                                 2) Capablility to load Drives's Vol. Label
'                                 3) New tree view line style
'                                 4) Will show it's own icons for *.ico, *.lnk, *.ani, *.cur
'                                       (in the prev ver, icons was for filetype)
'               Fixed the bug found, on expanding the last drive in the list...
'
'-----------------------------------------------------------------------------------------------------------

Option Explicit

'[APIs]
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function DrawTextA Lib "user32.dll" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextW Lib "user32.dll" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function DrawFocusRect Lib "user32.dll" (ByVal hDC As Long, ByRef lpRect As RECT) As Long
Private Declare Function Rectangle Lib "gdi32.dll" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer
Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (pDicDesc As TypeIcon, riid As CLSID, ByVal fown As Long, lpUnk As Object) As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function DrawIconEx Lib "user32.dll" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

' for subclassing
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long

' for API scroll bars
Private Declare Function InitialiseFlatSB Lib "comctl32.dll" Alias "InitializeFlatSB" (ByVal lhWnd As Long) As Long
Private Declare Function SetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal BOOL As Boolean) As Long
Private Declare Function GetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, LPSCROLLINFO As SCROLLINFO) As Long
Private Declare Function EnableScrollBar Lib "user32" (ByVal hwnd As Long, ByVal wSBflags As Long, ByVal wArrows As Long) As Long
Private Declare Function ShowScrollBar Lib "user32" (ByVal hwnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long
Private Declare Function FlatSB_EnableScrollBar Lib "comctl32.dll" (ByVal hwnd As Long, ByVal int2 As Long, ByVal UINT3 As Long) As Long
Private Declare Function FlatSB_ShowScrollBar Lib "comctl32.dll" (ByVal hwnd As Long, ByVal code As Long, ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_GetScrollInfo Lib "comctl32.dll" (ByVal hwnd As Long, ByVal code As Long, LPSCROLLINFO As SCROLLINFO) As Long
Private Declare Function FlatSB_SetScrollInfo Lib "comctl32.dll" (ByVal hwnd As Long, ByVal code As Long, LPSCROLLINFO As SCROLLINFO, ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_SetScrollProp Lib "comctl32.dll" (ByVal hwnd As Long, ByVal Index As Long, ByVal NewValue As Long, ByVal fRedraw As Boolean) As Long
Private Declare Function UninitializeFlatSB Lib "comctl32.dll" (ByVal hwnd As Long) As Long

'[Module Constants]
Private Const DIB_RGB_ColS      As Long = 0
Private Const VER_PLATFORM_WIN32_NT  As Long = 2
Private Const SHGFI_ICON = &H100
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1
Private Const SPLITER As String = "<%S%>"
Private Const GWL_EXSTYLE As Long = -20
Private Const WS_EX_TOOLWINDOW As Long = &H80&
Private Const SWP_SHOWWINDOW As Long = &H40

Private Const MAX_PATH = 260
Private Const INVALID_HANDLE_VALUE = -1

Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10

Private Const DRIVE_REMOVABLE = 2
Private Const DRIVE_FIXED = 3
Private Const DRIVE_REMOTE = 4
Private Const DRIVE_CDROM = 5
Private Const DRIVE_RAMDISK = 6

' for subclassing
Private Const WM_GETMINMAXINFO      As Long = &H24
Private Const WM_WINDOWPOSCHANGED   As Long = &H47
Private Const WM_WINDOWPOSCHANGING  As Long = &H46
Private Const WM_LBUTTONDOWN        As Long = &H201
Private Const WM_SIZE               As Long = &H5
Private Const WM_LBUTTONDBLCLK      As Long = &H203
Private Const WM_RBUTTONDOWN        As Long = &H204
Private Const WM_MOUSEMOVE          As Long = &H200
Private Const WM_SETFOCUS           As Long = &H7
Private Const WM_KILLFOCUS          As Long = &H8
Private Const WM_MOVE               As Long = &H3
Private Const WM_TIMER              As Long = &H113
Private Const WM_MOUSELEAVE         As Long = &H2A3
Private Const WM_MOUSEWHEEL         As Long = &H20A
Private Const WM_MOUSEHOVER         As Long = &H2A1

Private Const ALL_MESSAGES           As Long = -1                                       'All messages added or deleted
Private Const GMEM_FIXED             As Long = 0                                        'Fixed memory GlobalAlloc flag
Private Const GWL_WNDPROC            As Long = -4                                       'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04               As Long = 88                                       'Table B (before) address patch offset
Private Const PATCH_05               As Long = 93                                       'Table B (before) entry count patch offset
Private Const PATCH_08               As Long = 132                                      'Table A (after) address patch offset
Private Const PATCH_09               As Long = 137                                      'Table A (after) entry count patch offset

Private sc_aSubData()                As tSubData
Private bTrack                       As Boolean
Private bTrackUser32                 As Boolean
Private bInCtrl                      As Boolean

' for scroll bars
Private Const WM_VSCROLL = &H115
Private Const WM_HSCROLL = &H114

Private Const SB_BOTTOM = 7
Private Const SB_ENDSCROLL = 8
Private Const SB_HORZ = 0
Private Const SB_LEFT = 6
Private Const SB_LINEDOWN = 1
Private Const SB_LINELEFT = 0
Private Const SB_LINERIGHT = 1
Private Const SB_LINEUP = 0
Private Const SB_PAGEDOWN = 3
Private Const SB_PAGELEFT = 2
Private Const SB_PAGERIGHT = 3
Private Const SB_PAGEUP = 2
Private Const SB_RIGHT = 7
Private Const SB_THUMBTRACK = 5
Private Const SB_TOP = 6
Private Const SB_VERT = 1

Private Const SIF_RANGE = &H1
Private Const SIF_PAGE = &H2
Private Const SIF_POS = &H4
Private Const SIF_TRACKPOS = &H10
Private Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)

Private Const ESB_DISABLE_BOTH = &H3
Private Const ESB_ENABLE_BOTH = &H0
Private Const MK_CONTROL = &H8
Private Const WSB_PROP_VSTYLE = &H100&
Private Const WSB_PROP_HSTYLE = &H200&
Private Const FSB_FLAT_MODE = 1&
Private Const FSB_REGULAR_MODE = 0&

'[Types]
Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * 256
    szTypeName As String * 80
End Type

Private Type FILETIME
    dwLowDateTime     As Long
    dwHighDateTime    As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes  As Long
    ftCreationTime    As FILETIME
    ftLastAccessTime  As FILETIME
    ftLastWriteTime   As FILETIME
    nFileSizeHigh     As Long
    nFileSizeLow      As Long
    dwReserved0       As Long
    dwReserved1       As Long
    cFileName         As String * MAX_PATH
    cAlternate        As String * 14
End Type

Private Type TypeIcon
    cbSize As Long
    picType As PictureTypeConstants
    hIcon As Long
End Type

Private Type CLSID
    id((123)) As Byte
End Type

Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Type OSVERSIONINFO
   dwOSVersionInfoSize  As Long
   dwMajorVersion       As Long
   dwMinorVersion       As Long
   dwBuildNumber        As Long
   dwPlatformId         As Long
   szCSDVersion         As String * 128 ' Maintenance string
End Type

Private Type tSubData                                                                   'Subclass data type
    hwnd          As Long                                            'Handle of the window being subclassed
    nAddrSub      As Long                                            'The address of our new WndProc (allocated memory).
    nAddrOrig     As Long                                            'The address of the pre-existing WndProc
    nMsgCntA      As Long                                            'Msg after table entry count
    nMsgCntB      As Long                                            'Msg before table entry count
    aMsgTblA()    As Long                                            'Msg after table array
    aMsgTblB()    As Long                                            'Msg Before table array
End Type
                                
Private Type TRACKMOUSEEVENT_STRUCT
  cbSize          As Long
  dwFlags         As TRACKMOUSEEVENT_FLAGS
  hwndTrack       As Long
  dwHoverTime     As Long
End Type

' Scroll bar
Private Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type


'[Enums]
Public Enum TextAlignmentEnum
    [AL_Left] = 0
    [AL_Center] = 1
    [AL_Right] = 2
End Enum

Public Enum IconExtractEnum
    [SIZE_16] = 0
    [SIZE_32] = 1
End Enum

Public Enum ControlModeEnum
    [Mode_ListBox] = 0
    [Mode_DriveList] = 1
    [Mode_FolderList] = 2
    [Mode_FileList] = 3
    [Mode_FolderBrowser] = 4
    [Mode_FileBrowser] = 5
End Enum

Public Enum ListGradientDirectionEnum
    [Fill_None] = 0
    [Fill_Horizontal] = 1
    [Fill_HorizontalMiddleOut] = 2
    [Fill_Vertical] = 3
    [Fill_VerticalMiddleOut] = 4
    [Fill_DownwardDiagonal] = 5
    [Fill_UpwardDiagonal] = 6
End Enum

Public Enum List_AppearanceEnum
    [Flat] = 0
    [3D] = 1
End Enum

Public Enum List_BorderEnum
    BDR_None = 0
    BDR_RAISED = &H5
    BDR_StaticEdge = &H20000
    BDR_SUNKEN = &H200&
End Enum

Public Enum SortOrderEnum
    [Sort_None] = 0
    [Sort_Ascending] = -1
    [Sort_Desending] = 1
End Enum

Public Enum SelectionStyleEnum
    [Style_Normal] = 0
    [Style_XP] = 1
End Enum

' for subclassing
Private Enum eMsgWhen
    MSG_AFTER = 1                                                                         'Message calls back after the original (previous) WndProc
    MSG_BEFORE = 2                                                                        'Message calls back before the original (previous) WndProc
    MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE                                        'Message calls back before and after the original (previous) WndProc
End Enum

Private Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1&
    TME_LEAVE = &H2&
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum

' for scroll bars
Public Enum ScrollBarOrienationEnum
    Scroll_Horizontal
    Scroll_Vertical
    Scroll_Both
End Enum

Public Enum ScrollBarStyleEnum
    Style_Regular = FSB_REGULAR_MODE
    Style_Flat = FSB_FLAT_MODE
End Enum

Public Enum EFSScrollBarConstants
    efsHorizontal = SB_HORZ
    efsVertical = SB_VERT
End Enum

'[Local Variables]
Private m_SelItem       As Long
Private m_RowHeight     As Long
Private m_iCount        As Long
Private m_hMode         As Long
Private m_KeyControl    As Boolean
Private m_bIsNT         As Boolean
Private m_Inside        As Boolean
Private m_HasFocus      As Boolean
Private m_PicComplete   As PictureBox
Private m_Left          As Long
Private m_LastComplete  As Long
Private m_LeftShift     As Long
Private m_LastWidth     As Long

'[Data Storage]
Private m_Items         As New Collection
Private m_FileIcons     As New Collection
Private m_ImageList     As Object
Private m_Selected      As New Collection
Private m_TimerElsp     As Long

'[Property Variables]
Private m_Picture       As New StdPicture
Private m_BackColor     As OLE_COLOR
Private m_ForeColor     As OLE_COLOR
Private m_Font          As New StdFont
Private m_SelColor      As OLE_COLOR
Private m_FullRowSel    As Boolean
Private m_SortOrder     As SortOrderEnum
Private m_SelForeColor  As OLE_COLOR
Private m_StrechIcon    As Boolean
Private m_IconFocus     As Boolean
Private m_GridLines     As Boolean
Private m_GridColor     As OLE_COLOR
Private m_BackGradient  As ListGradientDirectionEnum
Private m_SelGradient   As ListGradientDirectionEnum
Private m_MultiSelect   As Boolean
Private m_SelStart      As Long
Private m_Mode          As ControlModeEnum
Private m_Path          As String
Private m_BackGradientCol   As OLE_COLOR
Private m_SelGradientCol    As OLE_COLOR
Private m_TextAlignment    As TextAlignmentEnum
Private m_ShowIcon          As Boolean
Private m_FocusRectangle    As Boolean
Private m_SelectionStyle    As SelectionStyleEnum
Private m_FlatScrollBar     As Boolean
Private m_AutoHideScrollBars As Boolean
Private m_BorderStyle       As List_BorderEnum
Private m_IconExtractSize   As Long
Private m_AutoRefresh       As Boolean
Private m_Filter            As String
Private m_ShowSystemFiles   As Boolean
Private m_ShowHiddenFiles   As Boolean

'[Default Property Values]
Private Const m_def_BackColor = &HFFFFFF
Private Const m_def_ForeColor = &H80000012
Private Const m_def_SelColor = &HFF8C1A
Private Const m_def_SelForeColor = &HFFFFFF
Private Const m_def_StrechIcon = False
Private Const m_def_Appearance = 1
Private Const m_def_BorderStyle = BDR_SUNKEN
Private Const m_def_FullRowSel = False
Private Const m_def_SortOrder = Sort_None
Private Const m_def_SortItems = False
Private Const m_def_IconFocus = True
Private Const m_def_TextAllignMent = vbLeftJustify
Private Const m_Def_RowHeight = 22
Private Const m_def_GridLines = False
Private Const m_Def_GridColor = &H8000000F
Private Const m_def_BackGradient = 0
Private Const m_def_SelGradient = 0
Private Const m_def_MultiSelect = False
Private Const m_def_SelCount = 0
Private Const m_def_BackGradientCol = &HC0C0C0
Private Const m_def_SelGradientCol = &HFFFFFF
Private Const m_def_ShowIcon = False
Private Const m_def_FocusRectangle = False
Private Const m_def_SelectionStyle = Style_XP
Private Const m_def_FlatScrollBar = False
Private Const m_def_AutoHideScrollBars = False
Private Const m_def_Mode = 0
Private Const m_def_Path = vbNullString
Private Const m_def_IconExtractSize = 0
Private Const m_def_AutoRefresh = True
Private Const m_def_Filter = "*.*"
Private Const m_def_ShowSystemFiles = False
Private Const m_def_ShowHiddenFiles = False

Private m_bInitialised      As Boolean
Private m_eOrientation      As ScrollBarOrienationEnum
Private m_eStyle            As ScrollBarStyleEnum
Private m_hWnd              As Long
Private m_lSmallChangeHorz  As Long
Private m_lSmallChangeVert  As Long
Private m_bEnabledHorz      As Boolean
Private m_bEnabledVert      As Boolean
Private m_bVisibleHorz      As Boolean
Private m_bVisibleVert      As Boolean
Private m_bNoFlatScrollBars As Boolean

'[Events]
Public Event MouseEnter()
Public Event MouseLeave()
Public Event Click()
Attribute Click.VB_MemberFlags = "200"
Public Event SelChange()
Public Event DbClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)


'[ Subclassed events receiver ]
'------------------------------------------------------------------------------------------
Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
Dim eBar As EFSScrollBarConstants
Dim lV As Long, lSC As Long
Dim lScrollCode As Long
Dim tSI As SCROLLINFO
Dim zDelta As Long
Dim lPrev_Vert As Long
Dim lPrev_Hor As Long

    lPrev_Vert = SBValue(efsVertical)
    lPrev_Vert = SBValue(efsHorizontal)
    
    ' Gray Nobles's API scrollbar code
    Select Case uMsg
    
        ' The timer callback for refreshing list
        Case WM_TIMER
            m_TimerElsp = m_TimerElsp + 1
            If m_TimerElsp = 5 Then ' After 1/2 Sec
                KillTimer hwnd, 1
                m_TimerElsp = 0
                Me.Refresh
            End If
            
        Case WM_VSCROLL, WM_HSCROLL, WM_MOUSEWHEEL
            lScrollCode = (wParam And &HFFFF&)

            Select Case uMsg
            
                Case WM_HSCROLL ' Get the scrollbar type
                    eBar = efsHorizontal
                    
                Case WM_VSCROLL
                    eBar = efsVertical
                    
                Case Else     'WM_MOUSEWHEEL
                    If m_Items.Count > m_iCount Then
                        eBar = IIf(lScrollCode And MK_CONTROL, efsHorizontal, efsVertical)
                        lScrollCode = IIf(wParam / 65536 < 0, SB_LINEDOWN, SB_LINEUP)
                    End If
            End Select

            Select Case lScrollCode
            
                Case SB_THUMBTRACK
                
                    ' Is vertical/horizontal?
                    pSBGetSI eBar, tSI, SIF_TRACKPOS
                    SBValue(eBar) = tSI.nTrackPos

                Case SB_LEFT, SB_BOTTOM
                     SBValue(eBar) = IIf(lScrollCode = 7, SBMax(eBar), SBMin(eBar))

                Case SB_RIGHT, SB_TOP
                     SBValue(eBar) = SBMin(eBar)

                Case SB_LINELEFT, SB_LINEUP
                
                    If SBVisible(eBar) Then
                    
                        lV = SBValue(eBar)
                        If (eBar = efsHorizontal) Then
                            lSC = m_lSmallChangeHorz
                        Else
                            lSC = m_lSmallChangeVert
                        End If
                        
                        If (lV - lSC < SBMin(eBar)) Then
                             SBValue(eBar) = SBMin(eBar)
                        Else
                             SBValue(eBar) = lV - lSC
                        End If
                        
                    End If

                Case SB_LINERIGHT, SB_LINEDOWN
                    If SBVisible(eBar) Then
            
                        lV = SBValue(eBar)
                        
                        If (eBar = efsHorizontal) Then
                            lSC = m_lSmallChangeHorz
                        Else
                            lSC = m_lSmallChangeVert
                        End If
                        
                        If (lV + lSC > SBMax(eBar)) Then
                             SBValue(eBar) = SBMax(eBar)
                        Else
                             SBValue(eBar) = lV + lSC
                        End If
                    End If

                Case SB_PAGELEFT, SB_PAGEUP
                     SBValue(eBar) = SBValue(eBar) - SBLargeChange(eBar)

                Case SB_PAGERIGHT, SB_PAGEDOWN
                     SBValue(eBar) = SBValue(eBar) + SBLargeChange(eBar)

                Case SB_ENDSCROLL

            End Select
            
            If Not lPrev_Vert = SBValue(efsVertical) Or Not lPrev_Hor = SBValue(efsHorizontal) Then ReDrawList
            
    End Select

    Select Case uMsg
    
        Case WM_MOUSEWHEEL
            m_PicComplete.Visible = False
            m_LastComplete = -100
            
        Case WM_MOUSEMOVE
            If Not bInCtrl Then
                bInCtrl = True
                Call TrackMouseLeave(lng_hWnd)
                RaiseEvent MouseEnter
            End If

        Case WM_MOUSELEAVE
            bInCtrl = False
            m_PicComplete.Visible = False
            m_LastComplete = -100
            RaiseEvent MouseLeave
            
    End Select


    
End Sub


'------------------------------------------------------------------------------------------
'Domain     : UserControl Properties
'------------------------------------------------------------------------------------------

Public Property Get Mode() As ControlModeEnum
    Mode = m_Mode
End Property

Public Property Let Mode(ByVal New_Mode As ControlModeEnum)

    If Not m_Mode = New_Mode Then
        m_Mode = New_Mode
        PropertyChanged "Mode"
        If New_Mode = Mode_ListBox Then
            Me.Clear
        Else
            LoadPath
        End If
        Me.Refresh
    End If
    
End Property

Public Property Get Path() As String
    Path = m_Path
End Property


Public Property Let Path(ByVal New_Path As String)
    If Not m_Path = New_Path And Not Mode = Mode_FileBrowser Then
        If Not Right(New_Path, 1) = "\" And Not New_Path = vbNullString Then New_Path = New_Path & "\"
        m_Path = New_Path
        PropertyChanged "Path"
        LoadPath
        Me.Refresh
    End If
End Property

Public Function ListCount() As Long
On Error GoTo handle
    ListCount = m_Items.Count
Exit Function
handle:
    ListCount = 0
End Function


Public Property Get ListIcon(ByVal Index As Long) As Long
    ListIcon = Split(m_Items(Index + 1), SPLITER)(1)
End Property

Public Property Let ListIcon(ByVal Index As Long, ByVal vNewPicture As Long)
Dim txtData() As String
    
    txtData() = Split(m_Items(Index + 1), SPLITER)
    txtData(1) = vNewPicture
    m_Items.Remove Index + 1
    m_Items.add Join(txtData, SPLITER), , Index + 1
    PropertyChanged "ListIcon"
    ReDrawList
    
End Property


Public Property Get ListBold(ByVal Index As Long) As Boolean
Dim mBold As String

    mBold = Split(m_Items(Index + 1), SPLITER)(2)
    If mBold = "True" Then
        ListBold = True
    Else
        ListBold = False
    End If
    
End Property

Public Property Let ListBold(ByVal Index As Long, ByVal vNewVlaue As Boolean)
Dim txtData() As String
    
    txtData = Split(m_Items(Index + 1), SPLITER)
    txtData(2) = vNewVlaue
    m_Items.Remove Index + 1
    
    If Index + 1 > m_Items.Count Then
        m_Items.add Join(txtData, SPLITER)
    Else
        m_Items.add Join(txtData, SPLITER), , Index + 1
    End If
    
    PropertyChanged "ListBold"
    ReDrawList

handle:
End Property


Public Property Get Picture() As Picture
    Set Picture = m_Picture
End Property

Public Property Set Picture(ByVal vNewPicture As Picture)
    Set m_Picture = vNewPicture
    PropertyChanged "Picture"
    ReDrawList
End Property


Public Property Get Font() As Font
    Set Font = m_Font
End Property

Public Property Set Font(ByVal vNewFont As Font)
    Set m_Font = vNewFont
    UserControl.Font = vNewFont
    If TextHeight("A") > m_RowHeight Then m_RowHeight = TextHeight("A")
    PropertyChanged "Font"
    Me.Refresh
End Property


Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal vNewCol As OLE_COLOR)
    m_ForeColor = vNewCol
    PropertyChanged "ForeColor"
    ReDrawList
End Property


Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal vNewCol As OLE_COLOR)
    m_BackColor = vNewCol
    PropertyChanged "BackColor"
    UserControl.BackColor = vNewCol
    ReDrawList
End Property


Public Property Get SelColor() As OLE_COLOR
    SelColor = m_SelColor
End Property

Public Property Let SelColor(ByVal vNewCol As OLE_COLOR)
    m_SelColor = vNewCol
    PropertyChanged "SelColor"
    ReDrawList
End Property


Public Property Get SelForeColor() As OLE_COLOR
    SelForeColor = m_SelForeColor
End Property

Public Property Let SelForeColor(ByVal vNewCol As OLE_COLOR)
    m_SelForeColor = vNewCol
    PropertyChanged "SelForeColor"
    ReDrawList
End Property


Public Property Get StrechIcon() As Boolean
    StrechIcon = m_StrechIcon
End Property

Public Property Let StrechIcon(ByVal vNewValue As Boolean)
    m_StrechIcon = vNewValue
    PropertyChanged "StrechIcon"
    ReDrawList
End Property


Public Property Get Appearance() As List_AppearanceEnum
    Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal vNewAppearance As List_AppearanceEnum)
    UserControl.Appearance = vNewAppearance
    PropertyChanged "Appearance"
    ReDrawList
End Property


Public Property Get BorderStyle() As List_BorderEnum
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal vNewBorder As List_BorderEnum)
    m_BorderStyle = vNewBorder
    UserControl.BorderStyle = 1
    SetWindowLong hwnd, GWL_EXSTYLE, m_BorderStyle
    SetWindowPos hwnd, 0, 0, 0, 0, 0, 55
    PropertyChanged "BorderStyle"
    ReDrawList
End Property


Public Property Get ListIndex() As Long
    If m_Items.Count = 0 Then
        ListIndex = -1
    Else
        ListIndex = m_SelItem
    End If
End Property

Public Property Let ListIndex(ByVal vNewValue As Long)
    
    If vNewValue < 0 Or vNewValue > m_Items.Count - 1 Then Exit Property
    m_SelItem = vNewValue
    
    If SBVisible(efsVertical) Then
        If m_SelItem < SBValue(efsVertical) Then SBValue(efsVertical) = IIf(SBValue(efsVertical) = 0, 0, m_SelItem)
        If m_SelItem >= SBValue(efsVertical) + Int(ScaleHeight / m_RowHeight) Then SBValue(efsVertical) = IIf(SBValue(efsVertical) = SBMax(efsVertical), SBMax(efsVertical), m_SelItem - m_iCount + 1)
    End If
    
    Selection_Clear
    m_Selected.add m_SelItem
    
    PropertyChanged "ListIndex"
    CheckSelected
    ReDrawList
    RaiseEvent SelChange
    
End Property


Public Property Get Text() As String
Attribute Text.VB_MemberFlags = "200"
    On Error GoTo handle
    If ListCount = 0 Then Exit Property
    Text = Split(m_Items(m_SelItem + 1), SPLITER)(0)
handle:
End Property


Public Property Get List(ByVal Index As Long) As String

    If Index > ListCount - 1 Or Index < 0 Then Exit Property
    List = Split(m_Items(Index + 1), SPLITER)(0)
    
End Property

Public Property Let List(ByVal Index As Long, ByVal vNewValue As String)
Dim txtData() As String
    
    txtData = Split(m_Items(Index + 1), SPLITER)
    txtData(0) = vNewValue
    m_Items(Index + 1) = Join(txtData, SPLITER)
    
End Property


Public Property Get FullRowSelect() As Boolean
    FullRowSelect = m_FullRowSel
End Property

Public Property Let FullRowSelect(ByVal vNewValue As Boolean)
    m_FullRowSel = vNewValue
    PropertyChanged "FullRowSelect"
    ReDrawList
End Property


Public Property Get MultiSelect() As Boolean
    MultiSelect = m_MultiSelect
End Property

Public Property Let MultiSelect(ByVal New_MultiSelect As Boolean)
    
    m_MultiSelect = New_MultiSelect
    Selection_Clear
    m_Selected.add m_SelItem
    
    PropertyChanged "MultiSelect"
    ReDrawList
    
End Property


Public Property Get SelCount() As Long
    SelCount = m_Selected.Count
End Property

Public Property Set ImageList(ByVal vNewValue As Object)
    Set m_ImageList = vNewValue
End Property


Public Property Get FocusRectangle() As Boolean
    FocusRectangle = m_FocusRectangle
End Property

Public Property Let FocusRectangle(ByVal New_FocusRectangle As Boolean)
    m_FocusRectangle = New_FocusRectangle
    PropertyChanged "FocusRectangle"
    ReDrawList
End Property


Public Property Get ShowIcon() As Boolean
    ShowIcon = m_ShowIcon
End Property

Public Property Let ShowIcon(ByVal New_ShowIcon As Boolean)
    m_ShowIcon = New_ShowIcon
    PropertyChanged "ShowIcon"
    ReDrawList
End Property


Public Property Get SortOrder() As SortOrderEnum
    SortOrder = m_SortOrder
End Property

Public Property Let SortOrder(ByVal vNewValue As SortOrderEnum)
    m_SortOrder = vNewValue
    PropertyChanged "SortOrder"
    SortCollection m_Items, m_SortOrder
    ReDrawList
End Property


Public Property Get IconFocus() As Boolean
    IconFocus = m_IconFocus
End Property

Public Property Let IconFocus(ByVal vNewValue As Boolean)
    m_IconFocus = vNewValue
    PropertyChanged "IconFocus"
    ReDrawList
End Property


Public Property Get TextAlignment() As TextAlignmentEnum
    TextAlignment = m_TextAlignment
End Property

Public Property Let TextAlignment(ByVal vNewValue As TextAlignmentEnum)
    m_TextAlignment = vNewValue
    PropertyChanged "TextAlignment"
    ReDrawList
End Property


Public Property Get GridLines() As Boolean
    GridLines = m_GridLines
End Property

Public Property Let GridLines(ByVal vNewValue As Boolean)
    m_GridLines = vNewValue
    PropertyChanged "GridLines"
    ReDrawList
End Property


Public Property Get GridColor() As OLE_COLOR
    GridColor = m_GridColor
End Property

Public Property Let GridColor(ByVal vNewValue As OLE_COLOR)
    m_GridColor = vNewValue
    PropertyChanged "GridColor"
    ReDrawList
End Property


Public Property Get RowHeight() As Long
    RowHeight = m_RowHeight
End Property

Public Property Let RowHeight(ByVal vNewValue As Long)
    If vNewValue >= TextHeight("A") Then
        m_RowHeight = vNewValue
    Else
        m_RowHeight = TextHeight("A")
    End If
    PropertyChanged "RowHeight"
    Me.Refresh
End Property


Public Property Get BackGradient() As ListGradientDirectionEnum
    BackGradient = m_BackGradient
End Property

Public Property Let BackGradient(ByVal New_BackGradient As ListGradientDirectionEnum)
    m_BackGradient = New_BackGradient
    PropertyChanged "BackGradient"
    ReDrawList
End Property


Public Property Get SelGradient() As ListGradientDirectionEnum
    SelGradient = m_SelGradient
End Property

Public Property Let SelGradient(ByVal New_SelGradient As ListGradientDirectionEnum)
    m_SelGradient = New_SelGradient
    PropertyChanged "SelGradient"
    ReDrawList
End Property

Public Property Get BackGradientCol() As OLE_COLOR
    BackGradientCol = m_BackGradientCol
End Property

Public Property Let BackGradientCol(ByVal New_BackGradientCol As OLE_COLOR)
    m_BackGradientCol = New_BackGradientCol
    PropertyChanged "BackGradientCol"
    ReDrawList
End Property


Public Property Get SelGradientCol() As OLE_COLOR
    SelGradientCol = m_SelGradientCol
End Property

Public Property Let SelGradientCol(ByVal New_SelGradientCol As OLE_COLOR)
    m_SelGradientCol = New_SelGradientCol
    PropertyChanged "SelGradientCol"
    ReDrawList
End Property


Public Property Get SelItem(ByVal vSelIndex As Long) As Long
    SelItem = m_Selected(vSelIndex + 1)
End Property


Public Property Get SelectionStyle() As SelectionStyleEnum
    SelectionStyle = m_SelectionStyle
End Property

Public Property Let SelectionStyle(ByVal vNewValue As SelectionStyleEnum)
    m_SelectionStyle = vNewValue
    PropertyChanged "SelectionStyle"
    ReDrawList
End Property


Public Property Get FlatScrollBar() As Boolean
    FlatScrollBar = m_FlatScrollBar
End Property

Public Property Let FlatScrollBar(ByVal vNewValue As Boolean)
    
    m_FlatScrollBar = vNewValue
    pSBClearUp
    m_bNoFlatScrollBars = Not m_FlatScrollBar
    If m_FlatScrollBar Then SBCreate UserControl.hwnd
    PropertyChanged "FlatScrollBar"
    DoEvents
    Me.Refresh
    
End Property


Public Property Get AutoHideScrollBars() As Boolean
    AutoHideScrollBars = m_AutoHideScrollBars
End Property

Public Property Let AutoHideScrollBars(ByVal New_AutoHideScrollBars As Boolean)
    m_AutoHideScrollBars = New_AutoHideScrollBars
    PropertyChanged "AutoHideScrollBars"
    Me.Refresh
End Property


Public Property Get IconExtractSize() As IconExtractEnum
    IconExtractSize = m_IconExtractSize
End Property

Public Property Let IconExtractSize(ByVal New_IconExtractSize As IconExtractEnum)
    m_IconExtractSize = New_IconExtractSize
    PropertyChanged "IconExtractSize"
    Set m_FileIcons = Nothing
    Set m_FileIcons = New Collection
    Me.Refresh
End Property


Public Property Get Filter() As String
    Filter = m_Filter
End Property

Public Property Let Filter(ByVal New_Filter As String)
    If Not New_Filter = m_Filter Then
        m_Filter = New_Filter
        LoadPath
        Me.Refresh
    End If
    PropertyChanged "Filter"
End Property

Public Property Get AutoRefresh() As Boolean
    AutoRefresh = m_AutoRefresh
End Property

Public Property Let AutoRefresh(ByVal New_AutoRefresh As Boolean)
    m_AutoRefresh = New_AutoRefresh
    PropertyChanged "AutoRefresh"
    If New_AutoRefresh = True Then Refresh
End Property


Public Property Get ShowSystemFiles() As Boolean
    ShowSystemFiles = m_ShowSystemFiles
End Property

Public Property Let ShowSystemFiles(ByVal New_ShowSystemFiles As Boolean)
    m_ShowSystemFiles = New_ShowSystemFiles
    PropertyChanged "ShowSystemFiles"
    LoadPath
    Me.Refresh
End Property


Public Property Get ShowHiddenFiles() As Boolean
    ShowHiddenFiles = m_ShowHiddenFiles
End Property

Public Property Let ShowHiddenFiles(ByVal New_ShowHiddenFiles As Boolean)
    m_ShowHiddenFiles = New_ShowHiddenFiles
    PropertyChanged "ShowHiddenFiles"
    LoadPath
    Me.Refresh
End Property

' By 'Gary Noble'
Public Sub ClearSelection()

    Selection_Clear
    m_Selected.add m_SelItem
    ReDrawList
    RaiseEvent SelChange
    
End Sub

' By 'Gary Noble'
Private Function FindSelection(ByVal vIdex As Long) As Long
On Error Resume Next

    Dim X As Long
    Dim xMax As Long

    xMax = m_Selected.Count

    For X = 1 To xMax
        If m_Selected(X) = vIdex Then FindSelection = X: Exit Function
    Next X
    
On Error GoTo 0
End Function

' By 'Gary Noble'
Public Sub SelectAll()

    On Error Resume Next

    Dim X As Long
    
    Selection_Clear
    For X = 0 To m_Items.Count
        m_Selected.add X
    Next
    ReDrawList
    RaiseEvent SelChange
    
    On Error GoTo 0
End Sub


'------------------------------------------------------------------------------------------
' Domain    : Events
'------------------------------------------------------------------------------------------

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
Dim X As Long
Dim lIndex As Long
Dim txtData() As String
Dim xMax As Long
Dim mFiles() As String
Dim mFolders() As String

    On Error GoTo handle
    If Mode = Mode_FolderBrowser Or Mode = Mode_FileBrowser Then
    
        txtData = Split(m_Items(m_SelItem + 1), SPLITER)
        txtData(0) = txtData(0) & "\"
        
        X = m_SelItem + 2
        If X <= m_Items.Count Then
            Do
                If InStr(1, m_Items(X), txtData(0)) = 1 Then
                    m_Items.Remove X
                    lIndex = 1
                Else
                    Exit Do
                End If
            Loop
        End If
        
        Me.ListBold(m_SelItem) = False
        If lIndex = 1 Then GoTo handle
        
        ScanPath txtData(0), mFiles, mFolders, "*.*"
        Me.ListBold(m_SelItem) = True
        lIndex = m_SelItem + 1
    
        xMax = GetMax(mFolders)
        If Not xMax = -1 Then
            For X = 0 To xMax
                If lIndex <= m_Items.Count - 1 Then
                    m_Items.add txtData(0) & mFolders(X) & SPLITER & "-1" & SPLITER & "False", , lIndex + 1
                Else
                    m_Items.add txtData(0) & mFolders(X) & SPLITER & "-1" & SPLITER & "False"
                End If
                lIndex = lIndex + 1
            Next X
        End If

        If m_Mode = Mode_FileBrowser Then
            Erase mFiles
            ScanPath txtData(0), mFiles, mFolders, m_Filter
            xMax = GetMax(mFiles)
            If Not xMax = -1 Then
                For X = 0 To xMax
                    If lIndex <= m_Items.Count - 1 Then
                        m_Items.add txtData(0) & mFiles(X) & SPLITER & "-1" & SPLITER & "False", , lIndex + 1
                    Else
                        m_Items.add txtData(0) & mFiles(X) & SPLITER & "-1" & SPLITER & "False"
                    End If
                    lIndex = lIndex + 1
                Next X
            End If
        End If
    End If
    
handle:
    Me.Refresh
    RaiseEvent DbClick
End Sub

Private Sub UserControl_GotFocus()
    m_HasFocus = True
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
Dim mNew As Long

    RaiseEvent KeyDown(KeyCode, Shift)
    
    ' Select each Key
    Select Case KeyCode
        Case vbKeyUp
            mNew = m_SelItem - 1
            
        Case vbKeyDown
            mNew = m_SelItem + 1
        
        Case vbKeyEnd
            mNew = ListCount - 1
        
        Case vbKeyHome
            mNew = 0
            
        Case vbKeyPageDown
            mNew = m_SelItem + m_iCount
        
        Case vbKeyPageUp
            mNew = m_SelItem - m_iCount
        
        Case Else
            Exit Sub
    End Select
    
    If mNew > ListCount - 1 Then mNew = ListCount - 1
    If mNew < 0 Then mNew = 0
            
    ' Refrech Control
    If Not mNew = m_SelItem And Not mNew = -1 Then

        If SBVisible(efsVertical) Then
            If mNew < SBValue(efsVertical) Then SBValue(efsVertical) = IIf(SBValue(efsVertical) = 0, 0, mNew)
            If mNew >= SBValue(efsVertical) + Int(ScaleHeight / m_RowHeight) Then SBValue(efsVertical) = IIf(SBValue(efsVertical) = SBMax(efsVertical), SBMax(efsVertical), mNew - m_iCount + 1)
        End If
        
        UserControl_MouseDown 999, 0, Val(mNew), 0
        
    End If

End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_LostFocus()
    m_HasFocus = False
    ReDrawList
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim mRowHeight  As Double
Dim xStart  As Long
Dim mIndex As Long
Dim xEnd  As Long
Dim X1 As Long

    If m_SelItem > ListCount - 1 Or m_SelItem < 0 Then Exit Sub
    
    
    If Button = 999 Then    ' Event send from KEyDown
        m_SelItem = X
    Else
        mRowHeight = (ScaleHeight / m_iCount)
        If Y / mRowHeight > m_Items.Count Then Exit Sub
        m_SelItem = SBValue(efsVertical) + Int(Y / mRowHeight)
    End If
    
    If GetKeyState(vbKeyControl) < 0 And Button = vbKeyLButton And m_MultiSelect Then
        mIndex = Selection_Find(m_SelItem)
        If mIndex = 0 Then
            m_Selected.add m_SelItem
            Selection_Sort
        Else
            m_Selected.Remove mIndex
        End If
        m_SelStart = m_SelItem
        GoTo Skip
    End If
    
    If GetKeyState(vbKeyShift) < 0 And m_MultiSelect Then
    
        If m_SelStart < m_SelItem Then
            xStart = m_SelStart
            xEnd = m_SelItem
        Else
            xStart = m_SelItem
            xEnd = m_SelStart
        End If
        
        Selection_Clear
        
        For X1 = xStart To xEnd
            m_Selected.add X1
        Next X1
        
        Selection_Sort
        
    Else
        m_SelStart = m_SelItem
        Selection_Clear
        m_Selected.add m_SelItem
        
    End If
    
    
Skip:
    DoEvents
    ReDrawList
    RaiseEvent SelChange
    RaiseEvent MouseDown(Button, Shift, X, Y)
    
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim mRowHeight  As Double
Static mPrev As Long
Dim xStart  As Long
Dim mIndex As Long
Dim xEnd  As Long
Dim X1 As Long

    If Y < 0 Or Y > ScaleHeight Then m_Inside = False Else m_Inside = True
    mRowHeight = (ScaleHeight / m_iCount)
    mIndex = Int(Y / mRowHeight)
    ShowItemComplete mIndex + 1

    ' In multiselection mode, user is trying to select more than one row
    ' by (LeftButton Down + MouseMove)
    If GetKeyState(vbKeyLButton) < 0 And m_MultiSelect And GetKeyState(vbKeyControl) >= 0 And GetKeyState(vbKeyShift) >= 0 Then

        m_SelItem = SBValue(efsVertical) + mIndex
        If Not IsValidSelection(m_SelItem) Then Exit Sub
        If m_SelItem = mPrev Then Exit Sub
        
        If SBVisible(efsVertical) Then
            If m_SelItem < SBValue(efsVertical) Then SBValue(efsVertical) = IIf(SBValue(efsVertical) = 0, 0, SBValue(efsVertical) - 1)
            If m_SelItem >= SBValue(efsVertical) + Int(ScaleHeight / m_RowHeight) Then SBValue(efsVertical) = IIf(SBValue(efsVertical) = SBMax(efsVertical), SBMax(efsVertical), SBValue(efsVertical) + 1)
        End If
        
        If m_SelStart < m_SelItem Then
            xStart = m_SelStart
            xEnd = m_SelItem
        Else
            xStart = m_SelItem
            xEnd = m_SelStart
        End If
        
        Selection_Clear
        For X1 = xStart To xEnd
            If X1 >= 0 Then m_Selected.add X1
        Next X1
        
        Selection_Sort
        ReDrawList
        If Not m_MultiSelect Then RaiseEvent SelChange
        
    End If
    
    ' Not in multiselect mode. The LeftButton is down and mouse is moving..
    ' So scroll down the list
    
NextSelection:

    If GetKeyState(vbKeyLButton) < 0 And Not m_MultiSelect Then
        
        mRowHeight = (ScaleHeight / m_iCount)
        m_SelItem = SBValue(efsVertical) + Int(Y / mRowHeight)
        If Not IsValidSelection(m_SelItem) Then Exit Sub
        
        If m_SelItem >= 0 And m_SelItem < m_Items.Count Then
        
            If Not m_SelItem = mPrev Then
            
                mPrev = m_SelItem
                Selection_Clear
                m_Selected.add m_SelItem
                
                If SBVisible(efsVertical) Then
                    If m_SelItem < SBValue(efsVertical) Then SBValue(efsVertical) = IIf(SBValue(efsVertical) = 0, 0, m_SelItem)
                    If m_SelItem >= SBValue(efsVertical) + Int(ScaleHeight / m_RowHeight) Then SBValue(efsVertical) = IIf(SBValue(efsVertical) = SBMax(efsVertical), SBMax(efsVertical), m_SelItem - m_iCount + 1)
                End If
                
                ReDrawList
                RaiseEvent SelChange
            End If
            
            DoEvents
            If Not m_Inside Then GoTo NextSelection
            
        End If
        
    End If

    If Int(Y / m_RowHeight) > Me.ListCount - 1 Then
        MousePointer = vbArrow
    Else
        MousePointer = vbCustom
    End If
    
    mPrev = m_SelItem
    RaiseEvent MouseMove(Button, Shift, X, Y)
    
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If m_SelItem > ListCount - 1 Or m_SelItem < 0 Then Exit Sub
    If m_MultiSelect Then RaiseEvent SelChange
    RaiseEvent MouseUp(Button, Shift, X, Y)
    
End Sub

Private Sub UserControl_Initialize()
    
    'Used to prevent crashes on XP
    
    m_hMode = LoadLibrary("shell32.dll")
    m_KeyControl = True
    
End Sub

Private Sub UserControl_InitProperties()

    
    
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_SelColor = m_def_SelColor
    m_SelForeColor = m_def_SelForeColor
    Set m_Picture = Nothing
    Set m_Font = Ambient.Font
    m_StrechIcon = m_def_StrechIcon
    m_RowHeight = m_Def_RowHeight
    m_FullRowSel = m_def_FullRowSel
    m_SortOrder = m_def_SortOrder
    m_IconFocus = m_def_IconFocus
    m_TextAlignment = m_def_TextAllignMent
    m_GridLines = m_def_GridLines
    m_GridColor = m_Def_GridColor
    
    m_BackGradient = m_def_BackGradient
    m_SelGradient = m_def_SelGradient
    m_BackGradientCol = m_def_BackGradientCol
    m_SelGradientCol = m_def_SelGradientCol
    m_MultiSelect = m_def_MultiSelect
    m_FocusRectangle = m_def_FocusRectangle
    m_ShowIcon = m_def_ShowIcon
    m_FlatScrollBar = m_def_FlatScrollBar
    m_bNoFlatScrollBars = Not m_def_FlatScrollBar
    m_AutoHideScrollBars = m_def_AutoHideScrollBars
    m_SelectionStyle = Style_XP
    Me.BorderStyle = m_def_BorderStyle
    Me.Appearance = m_def_Appearance
    Me.BackColor = vbWhite
    m_Mode = m_def_Mode
    m_Path = App.Path & "\"
    m_IconExtractSize = m_def_IconExtractSize
    
    
    If m_hWnd = 0 Then InitializeSubClassing
    

    m_Filter = m_def_Filter
    m_AutoRefresh = m_def_AutoRefresh
    m_ShowSystemFiles = m_def_ShowSystemFiles
    m_ShowHiddenFiles = m_def_ShowHiddenFiles
End Sub

Private Sub UserControl_Resize()
    
    Me.Refresh
End Sub

Private Sub UserControl_Terminate()
On Error GoTo Catch
   
    FreeLibrary m_hMode
    Me.Clear
    pSBClearUp
    Set m_Items = Nothing
    Set m_FileIcons = Nothing
    Call Subclass_StopAll

Catch:
End Sub

Private Sub CheckSelected()

    If m_Items.Count = 0 Then m_SelItem = -1
    If m_SelItem > m_Items.Count - 1 Then m_SelItem = m_Items.Count - 1
    If m_SelItem < 0 Then m_SelItem = 0

End Sub

'------------------------------------------------------------------------------------------
' Procedure  : AddItem
' Auther     : Jim Jose
' Input      : New item
' OutPut     : None
' Purpose    : To add an item to listBox
'------------------------------------------------------------------------------------------

Public Sub AddItem(Text As String, _
                    Optional Index As Long = -1, _
                    Optional Icon As Long = -1, _
                    Optional Bold As Boolean = False)
                
    If Not m_Mode = 0 Then Exit Sub
    
    If Index = -1 Then
        ' Index not specified , add to last
        m_Items.add Text & SPLITER & Icon & SPLITER & Bold
    Else
        ' add to specified index
        m_Items.add Text & SPLITER & Icon & SPLITER & Bold, , Index + 1
    End If
    
    If AutoRefresh Then
        SetTimer hwnd, 1, 1, 0
        m_TimerElsp = 0
    End If
    
End Sub

'------------------------------------------------------------------------------------------
' Procedure  : Remove
' Auther     : Jim Jose
' Input      : Index
' OutPut     : None
' Purpose    : To remove an item from List
'------------------------------------------------------------------------------------------

Public Sub Remove(Optional ByVal Index As Long = -1)
    
    If Not m_Mode = 0 Then Exit Sub
    
    If Index = -1 Then
        ' Index not specifid, remove selected item
        m_Items.Remove m_SelItem + 1
    Else
        ' Remove specified item
        m_Items.Remove Index + 1
    End If
    
    ' Sort If needed
    SortCollection m_Items, m_SortOrder
    Me.Refresh
    
End Sub

'------------------------------------------------------------------------------------------
' Procedure  : Clear
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : Clear List
'------------------------------------------------------------------------------------------

Public Sub Clear()
Dim X As Long

    If Not m_Mode = 0 Then Exit Sub
    
    ' Remove each Item
    For X = 1 To m_Items.Count
        m_Items.Remove (1)
    Next X
    Selection_Clear
    
    ' Redraw
    UserControl.Cls
    Me.Refresh
    
End Sub


'------------------------------------------------------------------------------------------
' Procedure  : Refresh
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : Arrage control and calculate local variables
'------------------------------------------------------------------------------------------

Public Sub Refresh()
On Error GoTo handle

    ' Determine item height & item cound per Screen
    m_iCount = Int(ScaleHeight / m_RowHeight)
    If m_iCount <= 0 Then m_iCount = 1
    
    ' Arrange\Set controls
    If m_Items.Count > m_iCount Then
        SBVisible(efsVertical) = True
        SBEnabled(efsVertical) = True
        SBMax(efsVertical) = m_Items.Count - m_iCount
        SBMax(efsVertical) = SBMax(efsVertical) - 6
    Else
        SBValue(efsVertical) = 0
        If m_AutoHideScrollBars Then
            SBVisible(efsVertical) = False
            SBEnabled(efsVertical) = True
        Else
            SBVisible(efsVertical) = True
            SBEnabled(efsVertical) = False
        End If
    End If
    SBSmallChange(efsVertical) = 1
    SBLargeChange(efsVertical) = 5
    SBSmallChange(efsHorizontal) = 5
    SBLargeChange(efsHorizontal) = 15
        
handle:

    SortCollection m_Items, m_SortOrder
    ReDrawList
    
End Sub


'------------------------------------------------------------------------------------------
' Procedure  : ReDrawList
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To draw the entire region
'------------------------------------------------------------------------------------------

Private Sub ReDrawList()

Dim mRowHeight  As Double
Dim mIcon       As StdPicture
Dim txtData()   As String
Dim mBold       As Boolean
Dim mText       As String
Dim mIndex      As Long
Dim xStart      As Long
Dim xMax        As Long
Dim mTop        As Long
Dim Rct         As RECT
Dim X           As Long
Dim TmpSelCol1  As Long
Dim TmpSelCol2  As Long
Dim TmpBorderCol As Long
Dim lPos        As Long
 
   
    On Error GoTo handle
    UserControl.Cls
    m_Left = 0
    CheckSelected
    xStart = SBValue(efsVertical)
    xMax = xStart + m_iCount
    If xMax > m_Items.Count - 1 Then xMax = m_Items.Count - 1
    mRowHeight = ScaleHeight / m_iCount
    UserControl.DrawStyle = vbTransparent
    If m_ShowIcon Then m_Left = mRowHeight + 3
    
    ' To view style on design mode
    If Ambient.UserMode = False And m_Items.Count = 0 And m_Mode = Mode_ListBox Then
        m_Items.add "McListBox" & SPLITER & "-1" & SPLITER & "False"
        m_Selected.add 0
    End If
    
    ' Draw back
    UserControl.Picture = m_Picture
    If m_BackGradient Then FillGradient hDC, 0, 0, ScaleWidth, ScaleHeight, m_BackColor, m_BackGradientCol, m_BackGradient
    
    ' Set selection colors
    If m_SelectionStyle = Style_XP Then
        TmpBorderCol = BlendColor(m_SelColor, vbBlack, 250)
        If m_HasFocus Then
            TmpSelCol1 = BlendColor(m_SelColor, vbWhite, 150)
            TmpSelCol2 = BlendColor(m_SelColor, vbWhite, 100)
        Else
            TmpSelCol1 = BlendColor(m_SelColor, vbWhite, 75)
            TmpSelCol2 = TmpSelCol1
        End If
    Else
        TmpBorderCol = m_SelColor
        TmpSelCol1 = m_SelColor
        TmpSelCol2 = TmpSelCol1
    End If
    
    m_LastWidth = 0
    For X = xStart To xMax

        ' Get the data
        txtData = Split(m_Items(X + 1), SPLITER)
        mText = txtData(0)
        
        ' Find if Selected!?
        mIndex = Selection_Find(X)
        mTop = (X - xStart) * mRowHeight '
        
        m_LeftShift = -SBValue(efsHorizontal)
        If Mode = Mode_FileBrowser Or Mode = Mode_FolderBrowser Then CheckItem mText, m_LeftShift, mTop
        
        'Load Icon
        If m_ShowIcon Then
            If Not m_Mode = Mode_ListBox Then
            
                'Hybrid Modes! Load icons from file itself
                Select Case Mode
                    Case Mode_FileBrowser, Mode_FolderBrowser
                         Set mIcon = GetIcon(txtData(0))
                    Case Mode_DriveList
                        Set mIcon = GetIcon(mText)
                    Case Else
                        Set mIcon = GetIcon(m_Path & mText)
                End Select
                
            Else
                ' Listbox Mode! Load icon from McImageList
                If Not txtData(1) = -1 Then Set mIcon = m_ImageList.ListImages(Val(txtData(1)))
            End If
        End If
        
        'Item bold
        If txtData(2) = "True" Then mBold = True Else mBold = False
        
        Select Case m_Mode
            Case Mode_DriveList, Mode_FileBrowser, Mode_FolderBrowser
                If m_LeftShift + SBValue(efsHorizontal) = 0 Then mText = GetVolumeLabel(mText & "\") & " (" & mText & ")"
        End Select
            
            
        If mIndex > 0 Then

            ' Fill Selection
            If m_SelGradient = Fill_None Then

                UserControl.ForeColor = TmpBorderCol
                If X = m_SelItem Then UserControl.FillColor = TmpSelCol1 Else UserControl.FillColor = TmpSelCol2

                If m_FullRowSel Then
                    Rectangle hDC, m_LeftShift + 1, mTop + 1, ScaleWidth, mTop + mRowHeight
                Else
                    Rectangle hDC, m_LeftShift + m_Left + 1, mTop + 1, ScaleWidth, mTop + mRowHeight
                End If

            Else
                If m_FullRowSel Then
                    FillGradient hDC, m_LeftShift + 0, mTop + 1, ScaleWidth, mRowHeight, m_SelColor, m_SelGradientCol, m_SelGradient
                Else
                    FillGradient hDC, m_LeftShift + m_Left, mTop + 1, ScaleWidth - m_Left, mRowHeight, m_SelColor, m_SelGradientCol, m_SelGradient
                End If

            End If

            ' Draw Icon Focus
            If (m_IconFocus And Not m_FullRowSel And m_ShowIcon And m_HasFocus) Then
                UserControl.ForeColor = vbBlack
                SetRect Rct, m_LeftShift + 1, mTop + 1, mRowHeight + m_LeftShift, mTop + mRowHeight
                DrawFocusRect UserControl.hDC, Rct
            End If

            UserControl.ForeColor = m_SelForeColor

        Else
        
            If (Mode = Mode_FileBrowser Or Mode = Mode_FolderBrowser) And m_IconFocus Then
                UserControl.ForeColor = vbGrayText
                UserControl.FillColor = m_BackColor
                Rectangle hDC, m_LeftShift + 1, mTop + 1, mRowHeight + m_LeftShift, mTop + mRowHeight
            End If
            
            UserControl.ForeColor = m_ForeColor
        End If
        
        ' Draw the Text
        UserControl.FontBold = mBold
        DrawText mText, m_LeftShift + m_Left, mTop + 1
        If (m_LeftShift + TextWidth(mText)) > m_LastWidth Then m_LastWidth = m_LeftShift + TextWidth(mText)
        
        ' Draw Icon
        If m_Mode = Mode_ListBox Then
            If Not txtData(1) = -1 And m_ShowIcon Then DrawPicture mIcon, m_LeftShift + 2, mTop + 2, mRowHeight - 4, mRowHeight - 4
        Else
            If m_ShowIcon Then DrawPicture mIcon, m_LeftShift + 2, mTop + 2, mRowHeight - 4, mRowHeight - 4
        End If

        'Draw Grid
        If m_GridLines Then
            UserControl.ForeColor = m_GridColor
            Line (0, mTop)-(ScaleWidth, mTop)
        End If
        
        If X = m_SelItem Then
        
            ' Draw Focus Rect
            If m_FocusRectangle Then
                UserControl.ForeColor = vbBlack
                If m_FullRowSel Or Not m_ShowIcon Then
                    SetRect Rct, m_LeftShift + -1, mTop, ScaleWidth + 1, mTop + mRowHeight + 1
                Else
                    SetRect Rct, m_LeftShift + m_Left, mTop, ScaleWidth + 1, mTop + mRowHeight + 1
                End If
                DrawFocusRect hDC, Rct
            End If
            
        End If
    Next X
    
    If m_GridLines Then Line (0, mTop + m_RowHeight)-(ScaleWidth, mTop + m_RowHeight)
    m_LastWidth = m_LastWidth + SBValue(efsHorizontal)
    
    If m_LastWidth > ScaleWidth Then
        SBVisible(efsHorizontal) = True
        SBEnabled(efsHorizontal) = True
        SBMax(efsHorizontal) = 50 + m_LastWidth - ScaleWidth
    Else
        SBValue(efsHorizontal) = 0
        If m_AutoHideScrollBars Then
            SBVisible(efsHorizontal) = False
            SBEnabled(efsHorizontal) = True
        Else
            SBVisible(efsHorizontal) = True
            SBEnabled(efsHorizontal) = False
        End If
    End If
    
    UserControl.Refresh
   
    
handle:
End Sub



Private Sub LoadPath()
Dim X As Long
Dim xMax As Long
Dim mDrives() As String
Dim mFiles() As String
Dim mFolders() As String

    ' Remove each Item
    Set m_Items = Nothing
    Set m_Items = New Collection
    Set m_FileIcons = Nothing
    Set m_FileIcons = New Collection
    Selection_Clear
    
    On Error GoTo handle
    Select Case m_Mode
        Case Mode_DriveList, Mode_FileBrowser, Mode_FolderBrowser
            mDrives = GetDrives
            xMax = UBound(mDrives)
            For X = 0 To xMax
                m_Items.add mDrives(X) & SPLITER & "-1" & SPLITER & "False"
            Next X
            
        Case Mode_FolderList
            ScanPath m_Path, mFiles, mFolders, "*.*"
            xMax = UBound(mFolders)
            For X = 0 To xMax
                m_Items.add mFolders(X) & SPLITER & "-1" & SPLITER & "False"
            Next X
            
        Case Mode_FileList
            ScanPath m_Path, mFiles, mFolders, m_Filter
            xMax = UBound(mFiles)
            For X = 0 To xMax
                m_Items.add mFiles(X) & SPLITER & "-1" & SPLITER & "False"
            Next X
            
    End Select
    
handle:
End Sub


Public Sub CheckItem(sText As String, lShift As Long, lTop As Long)
 
 Dim lPos As Long
 
    Do
        lPos = InStr(lPos + 1, sText, "\")
        If lPos = 0 Then
            lPos = InStrRev(sText, "\")
            sText = Right(sText, Len(sText) - lPos)
            If Not lTop = -1 Then Line (lShift - m_RowHeight / 2, lTop + m_RowHeight / 2)-(lShift, lTop + m_RowHeight / 2), vbGrayText
            Exit Do
        End If

        If Not lTop = -1 Then Line (lShift + m_RowHeight / 2, lTop)-(lShift + m_RowHeight / 2, lTop + m_RowHeight + 2), vbGrayText
        lShift = lShift + m_RowHeight

    Loop
    
End Sub


Public Function GetMax(sArr() As String) As Long
On Error GoTo handle
    GetMax = UBound(sArr)
Exit Function
handle:
GetMax = -1
End Function


Private Function GetIcon(ByVal sFileName As String) As StdPicture
Dim lPos As Long
Dim sKey As String

    On Error GoTo handle
    'This is a tough subject!! We have to store the icons loaded by 'GetFileIcon' function
    'for better speed. But storing all the icons is carzy when the filecount become 100 or more...
    'So we are storing the icons for each filetype(based on extesions)
    
    'BUT the problem is not ended... exe's will have different ICONS (ie we need to store icons
    'for each files there)
    
    If Mode = Mode_DriveList Then
        sKey = sFileName
    Else
        lPos = InStrRev(sFileName, ".")
        If lPos = 0 Then
            If InStr(1, sFileName, ":") = Len(sFileName) Then   'is drive
                sKey = sFileName
            Else
                sKey = "default"    'files with no ext
            End If
        Else
            sKey = StrConv(Right(sFileName, Len(sFileName) - lPos), vbLowerCase) 'files
        End If
    End If
    
    Select Case sKey
        Case "exe", "lnk", "ico", "ani", "Cur"
            Set GetIcon = GetFileIcon(sFileName, m_IconExtractSize)
        Case Else
            Set GetIcon = m_FileIcons(sKey)
    End Select
    
Exit Function
handle:
    m_FileIcons.add GetFileIcon(sFileName, m_IconExtractSize), sKey
    Set GetIcon = m_FileIcons(sKey)
    
    
End Function


Private Sub DrawText(ByVal lpStr As String, _
                        ByVal X As Long, ByVal Y As Long)
Dim Rct As RECT

    ' Set the Rect
    Rct.Left = X + 5
    Rct.Top = Y + (m_RowHeight - TextHeight("A")) / 2
    Rct.Right = ScaleWidth
    Rct.Bottom = Y + m_RowHeight
    
    ' Draw the text
    If IsNT Then
       DrawTextW hDC, StrPtr(lpStr), -1, Rct, m_TextAlignment
    Else
       DrawTextA hDC, lpStr, -1, Rct, m_TextAlignment
    End If
    
End Sub


Private Sub DrawPicture(mPicture As StdPicture, _
                        ByVal X As Long, ByVal Y As Long, _
                        ByVal lWidth As Long, ByVal lHeight As Long)
 Dim picWidth As Long
 Dim picHeight As Long
 
    If Not m_StrechIcon Then
        picWidth = ScaleX(mPicture.Width)
        picHeight = ScaleY(mPicture.Height)
        X = X + (lWidth - picWidth) / 2
        Y = Y + (lHeight - picHeight) / 2
    Else
        picWidth = lWidth
        picHeight = lHeight
    End If
        
    If mPicture.Type = vbPicTypeIcon Then
        Call DrawIconEx(hDC, X, Y, mPicture.handle, picWidth, picHeight, 0, 0, &H3)
    Else
        PaintPicture mPicture, X, Y, picWidth, picHeight
    End If
        
End Sub

Private Sub InitializeSubClassing()
On Error GoTo handle

    If Not m_hWnd = 0 Then Exit Sub
    
    ' Subclass in runtime
    If Ambient.UserMode Then
    
    bTrack = True
    bTrackUser32 = IsFunctionExported("TrackMouseEvent", "User32")
  
    If Not bTrackUser32 Then
      If Not IsFunctionExported("_TrackMouseEvent", "Comctl32") Then
        bTrack = False
      End If
    End If
    
    If Not bTrack Then Exit Sub
    
        With UserControl
            
            ' Start subclassing our calendar
            Call Subclass_Start(.hwnd)
            
            ' Adding the messages we need to track
            Call Subclass_AddMsg(.hwnd, WM_MOUSEWHEEL, MSG_AFTER)
            Call Subclass_AddMsg(.hwnd, WM_MOUSEMOVE, MSG_AFTER)
            Call Subclass_AddMsg(.hwnd, WM_MOUSELEAVE, MSG_AFTER)
            Call Subclass_AddMsg(.hwnd, WM_MOUSEHOVER, MSG_AFTER)
            Call Subclass_AddMsg(.hwnd, WM_VSCROLL, MSG_AFTER)
            Call Subclass_AddMsg(.hwnd, WM_HSCROLL, MSG_AFTER)
            Call Subclass_AddMsg(.hwnd, WM_TIMER, MSG_AFTER)
            
        End With
    
    End If
    
    'Create Standard Scrollbars
    If m_hWnd = 0 Then SBCreate UserControl.hwnd
    CreateItemComplete
    
handle:
End Sub

Private Sub Selection_Clear()
    Set m_Selected = Nothing
    Set m_Selected = New Collection
End Sub


Private Function Selection_Find(ByVal vIdex As Long) As Long
Dim X As Long
Dim xMax As Long

    xMax = m_Selected.Count
    
    For X = 1 To xMax
        If m_Selected(X) = vIdex Then Selection_Find = X: Exit Function
    Next X
    
End Function


Private Sub Selection_Sort()
Dim X As Long
Dim vPos As Long
Dim vCount As Long
Dim vStart As Long
Dim vNewCount As Long
Dim vNew As New Collection

    vStart = 1
    vCount = m_Selected.Count
    
    For X = vStart To vCount
        vNewCount = vNew.Count
        For vPos = 1 To vNewCount
            If vNew(vPos) < m_Selected(X) Then Exit For
        Next vPos
        
        If vPos = 1 Then
            vNew.add m_Selected(X)
        Else
            vNew.add m_Selected(X), , vPos - 1
        End If
        
    Next X
    
    Set m_Selected = vNew
    
End Sub


Private Sub CreateItemComplete()

    ' Create a new object
    Set m_PicComplete = UserControl.Controls.add("vb.PictureBox", "PicComplete")
    m_PicComplete.AutoRedraw = True
    m_PicComplete.BorderStyle = 1
    m_PicComplete.Appearance = 0
    m_PicComplete.BackColor = UserControl.FillColor
    m_PicComplete.ForeColor = vbBlack
    m_PicComplete.Enabled = False

    ' Hide window from TaskBar
    SetParent m_PicComplete.hwnd, GetDesktopWindow
    SetWindowLong m_PicComplete.hwnd, GWL_EXSTYLE, WS_EX_TOOLWINDOW
    
End Sub


Private Sub ShowItemComplete(ByVal lItem As Long)
Dim mArray() As String
Dim mTmp As Long
Dim Rct As RECT
Dim wRect As RECT
Dim ScrX As Long
Dim ScrY As Long
Dim lWidth As Long
Dim sText As String
Dim mRowHeight As Double

    On Error Resume Next
    mArray = Split(m_Items(SBValue(efsVertical) + lItem), SPLITER)
    sText = mArray(0)
    m_LeftShift = -SBValue(efsHorizontal)
    CheckItem sText, m_LeftShift, -1
    
    lWidth = TextWidth(sText)
    If lWidth < ScaleWidth - m_Left - m_LeftShift Then GoTo handle
    UserControl.FontBold = mArray(2)
    If m_LastComplete = lItem Then Exit Sub
    mRowHeight = ScaleHeight / m_iCount
    ScrX = Screen.TwipsPerPixelX
    ScrY = Screen.TwipsPerPixelY
    
    ' Set the Rect
    Rct.Left = 5
    Rct.Top = (mRowHeight - TextHeight("A")) / 2 - 2
    Set m_PicComplete.Font = m_Font
    m_PicComplete.Cls
    
    ' Move it
    GetWindowRect hwnd, wRect
    mTmp = Screen.Width / ScrX - (wRect.Left + m_Left)
    If mTmp < lWidth Then
    
        lWidth = mTmp - 10
        mArray = SplitToLines(sText, mTmp - 15)
        sText = Join(mArray, vbCrLf)
        mTmp = TextHeight(sText)
        If mTmp < mRowHeight Then
            Rct.Top = (mRowHeight - mTmp) / 2 - 2
        Else
            Rct.Top = 5
        End If
        m_PicComplete.Move (wRect.Left + m_Left + 2 + m_LeftShift) * ScrX, Int((wRect.Top + 3 + (lItem - 1) * mRowHeight) * ScrY), (lWidth) * ScrX, (5 + mTmp + 5) * ScrY
    
    Else
        m_PicComplete.Move (wRect.Left + m_Left + 2 + m_LeftShift) * ScrX, Int((wRect.Top + 3 + (lItem - 1) * mRowHeight) * ScrY), (lWidth + 12) * ScrX, (mRowHeight - 1) * ScrY
    End If
    
    Rct.Right = m_PicComplete.ScaleWidth
    Rct.Bottom = m_PicComplete.ScaleHeight
    
    ' Draw the text
    If IsNT Then
       DrawTextW m_PicComplete.hDC, StrPtr(sText), -1, Rct, vbLeftJustify
    Else
       DrawTextA m_PicComplete.hDC, sText, -1, Rct, vbLeftJustify
    End If

    ' Show
    m_PicComplete.ZOrder (0)
    m_PicComplete.Visible = True
    m_LastComplete = lItem
    
Exit Sub
handle:
    m_PicComplete.Visible = False
    m_LastComplete = -100
End Sub


' By Gary Noble
Private Function IsValidSelection(ByVal lSelectedItemIndex As Long) As Boolean
On Error Resume Next

        If lSelectedItemIndex > ListCount Then
            m_SelItem = ListCount - 1
            IsValidSelection = False
        ElseIf lSelectedItemIndex <= -1 Then
            m_SelItem = 0
            IsValidSelection = False
        Else
            m_SelItem = lSelectedItemIndex
            IsValidSelection = True
        End If

On Error GoTo 0
End Function

'------------------------------------------------------------------------------------------
' Procedure  : UserControl_WriteProperties
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : Write design time propery changes
'------------------------------------------------------------------------------------------

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    
    Call PropBag.WriteProperty("Picture", m_Picture, Nothing)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("SelColor", m_SelColor, m_def_SelColor)
    Call PropBag.WriteProperty("SelForeColor", m_SelForeColor, m_def_SelForeColor)
    Call PropBag.WriteProperty("StrechIcon", m_StrechIcon, m_def_StrechIcon)
    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, m_def_Appearance)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("FullRowSelect", m_FullRowSel, m_def_FullRowSel)
    Call PropBag.WriteProperty("SortOrder", m_SortOrder, m_def_SortOrder)
    Call PropBag.WriteProperty("IconFocus", m_IconFocus, m_def_IconFocus)
    Call PropBag.WriteProperty("TextAlignment", m_TextAlignment, m_def_TextAllignMent)
    Call PropBag.WriteProperty("RowHeight", m_RowHeight, m_Def_RowHeight)
    Call PropBag.WriteProperty("GridLines", m_GridLines, m_def_GridLines)
    Call PropBag.WriteProperty("GridColor", m_GridColor, m_Def_GridColor)
    Call PropBag.WriteProperty("BackGradient", m_BackGradient, m_def_BackGradient)
    Call PropBag.WriteProperty("SelGradient", m_SelGradient, m_def_SelGradient)
    Call PropBag.WriteProperty("BackGradientCol", m_BackGradientCol, m_def_BackGradientCol)
    Call PropBag.WriteProperty("SelGradientCol", m_SelGradientCol, m_def_SelGradientCol)
    Call PropBag.WriteProperty("MultiSelect", m_MultiSelect, m_def_MultiSelect)
    Call PropBag.WriteProperty("FocusRectangle", m_FocusRectangle, m_def_FocusRectangle)
    Call PropBag.WriteProperty("ShowIcon", m_ShowIcon, m_def_ShowIcon)
    Call PropBag.WriteProperty("SelectionStyle", m_SelectionStyle, m_def_SelectionStyle)
    Call PropBag.WriteProperty("FlatScrollBar", m_FlatScrollBar, m_def_FlatScrollBar)
    Call PropBag.WriteProperty("AutoHideScrollBars", m_AutoHideScrollBars, m_def_AutoHideScrollBars)
    Call PropBag.WriteProperty("Mode", m_Mode, m_def_Mode)
    Call PropBag.WriteProperty("Path", m_Path, m_def_Path)
    Call PropBag.WriteProperty("IconExtractSize", m_IconExtractSize, m_def_IconExtractSize)
    Call PropBag.WriteProperty("Filter", m_Filter, m_def_Filter)
    Call PropBag.WriteProperty("AutoRefresh", m_AutoRefresh, m_def_AutoRefresh)
    Call PropBag.WriteProperty("ShowSystemFiles", m_ShowSystemFiles, m_def_ShowSystemFiles)
    Call PropBag.WriteProperty("ShowHiddenFiles", m_ShowHiddenFiles, m_def_ShowHiddenFiles)
    

End Sub

'------------------------------------------------------------------------------------------
' Procedure  : UserControl_ReadProperties
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : Read design time propery changes
'------------------------------------------------------------------------------------------

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    
    m_SelItem = -1
    Set m_Picture = PropBag.ReadProperty("Picture", Nothing)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_SelColor = PropBag.ReadProperty("SelColor", m_def_SelColor)
    m_SelForeColor = PropBag.ReadProperty("SelForeColor", m_def_SelForeColor)
    m_StrechIcon = PropBag.ReadProperty("StrechIcon", m_def_StrechIcon)
    Me.Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
    Me.BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_FullRowSel = PropBag.ReadProperty("FullRowSelect", m_def_FullRowSel)
    m_SortOrder = PropBag.ReadProperty("SortOrder", m_def_SortOrder)
    m_IconFocus = PropBag.ReadProperty("IconFocus", m_def_IconFocus)
    m_TextAlignment = PropBag.ReadProperty("TextAlignment", m_def_TextAllignMent)
    m_RowHeight = PropBag.ReadProperty("RowHeight", m_Def_RowHeight)
    Set UserControl.Font = m_Font
    If m_RowHeight < TextHeight("A") Then m_RowHeight = TextHeight("A")
    m_GridLines = PropBag.ReadProperty("GridLines", m_def_GridLines)
    m_GridColor = PropBag.ReadProperty("GridColor", m_Def_GridColor)
    m_BackGradient = PropBag.ReadProperty("BackGradient", m_def_BackGradient)
    m_SelGradient = PropBag.ReadProperty("SelGradient", m_def_SelGradient)
    m_BackGradientCol = PropBag.ReadProperty("BackGradientCol", m_def_BackGradientCol)
    m_SelGradientCol = PropBag.ReadProperty("SelGradientCol", m_def_SelGradientCol)
    m_MultiSelect = PropBag.ReadProperty("MultiSelect", m_def_MultiSelect)
    UserControl.BackColor = m_BackColor
    m_FocusRectangle = PropBag.ReadProperty("FocusRectangle", m_def_FocusRectangle)
    m_ShowIcon = PropBag.ReadProperty("ShowIcon", m_def_ShowIcon)
    m_SelectionStyle = PropBag.ReadProperty("SelectionStyle", m_def_SelectionStyle)
    m_AutoHideScrollBars = PropBag.ReadProperty("AutoHideScrollBars", m_def_AutoHideScrollBars)
    m_FlatScrollBar = PropBag.ReadProperty("FlatScrollBar", m_def_FlatScrollBar)
    m_bNoFlatScrollBars = Not m_FlatScrollBar
    m_IconExtractSize = PropBag.ReadProperty("IconExtractSize", m_def_IconExtractSize)
    m_Mode = PropBag.ReadProperty("Mode", m_def_Mode)
    m_Path = PropBag.ReadProperty("Path", m_def_Path)
    m_Filter = PropBag.ReadProperty("Filter", m_def_Filter)
    m_AutoRefresh = PropBag.ReadProperty("AutoRefresh", m_def_AutoRefresh)
    m_ShowSystemFiles = PropBag.ReadProperty("ShowSystemFiles", m_def_ShowSystemFiles)
    m_ShowHiddenFiles = PropBag.ReadProperty("ShowHiddenFiles", m_def_ShowHiddenFiles)
    
    InitializeSubClassing
    
    LoadPath
    Me.Refresh


End Sub

'------------------------------------------------------------------------------------------
' Procedure  : GetDrives
' Auther     : Jim Jose (Thanks to Peter van Vessem)
' Input      : None!
' OutPut     : DRives
' Purpose    : To get the DriveNames on PC.
'------------------------------------------------------------------------------------------

Private Function GetDrives() As String()
Dim mArray() As String
Dim mFsObj As Object
Dim mDrv As Object
Dim X As Long

    Set mFsObj = CreateObject("scripting.filesystemobject")
    ReDim mArray(mFsObj.Drives.Count - 1)
    
    For Each mDrv In mFsObj.Drives
        mArray(X) = mDrv
        X = X + 1
    Next
    GetDrives = mArray
    
End Function


'------------------------------------------------------------------------------------------
' Procedure  : ScanPath
' Auther     : Jim Jose (Thanks to Richard Mewett)
' Input      : Path
' OutPut     : Files+Folders
' Purpose    : Get the Files and folders in the input path/folder!
'------------------------------------------------------------------------------------------

Private Function ScanPath(ByVal sPath As String, _
                    ByRef SFiles() As String, _
                    ByRef sFolders() As String, _
                    Optional ByVal sFilter As String = "*.*") As Boolean

 Dim lResult    As Long
 Dim mWFD       As WIN32_FIND_DATA
 Dim TmpName    As String
 Dim lFiles     As Long
 Dim lFolders   As Long
 Dim lFlag1      As Long
 Dim lFlag2      As Long
 
    On Error GoTo handle
    If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
    lResult = FindFirstFile(sPath & sFilter, mWFD)
    If lResult = INVALID_HANDLE_VALUE Then ScanPath = False: Exit Function

    If Not m_ShowHiddenFiles Then lFlag1 = FILE_ATTRIBUTE_HIDDEN
    If Not m_ShowSystemFiles Then lFlag2 = FILE_ATTRIBUTE_SYSTEM

    Do
        TmpName = Left$(mWFD.cFileName, InStr(mWFD.cFileName, Chr$(0)) - 1)
   
        Select Case TmpName
            Case ".", ".."
            
            Case Else
                If (mWFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
                
                    If (mWFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) And (mWFD.dwFileAttributes And lFlag2) = False And (mWFD.dwFileAttributes And lFlag1) = False Then
                        ReDim Preserve sFolders(lFolders)
                        sFolders(lFolders) = TmpName
                        lFolders = lFolders + 1
                    End If
                    
                ElseIf (mWFD.dwFileAttributes And lFlag2) = False And (mWFD.dwFileAttributes And lFlag1) = False Then
                    ReDim Preserve SFiles(lFiles)
                    SFiles(lFiles) = TmpName
                    lFiles = lFiles + 1
                End If
                   
        End Select

    Loop While FindNextFile(lResult, mWFD)
    
    ScanPath = True
    
Exit Function
handle:
ScanPath = False
End Function



'------------------------------------------------------------------------------------------
' Procedure  : GetFileIcon
' Auther     : Jim Jose (Thanks to D. Rijmenants)
' Input      : FIleName
' OutPut     : FileIcon
' Purpose    : To extract a the display Icon from any file or folder!
'------------------------------------------------------------------------------------------

Private Function GetFileIcon(ByVal sFileName As String, _
                                ByVal IconSize As IconExtractEnum) As StdPicture
 Dim SHinfo As SHFILEINFO
 Dim mTYPEICON As TypeIcon
 Dim mCLSID As CLSID
 Dim hIcon As Long
 Dim lFlag As Long

    If IconSize = [SIZE_16] Then lFlag = SHGFI_SMALLICON Else lFlag = SHGFI_LARGEICON
    If Right(sFileName, 1) <> "\" Then sFileName = sFileName & "\"
    Call SHGetFileInfo(sFileName, 0, SHinfo, Len(SHinfo), SHGFI_ICON + lFlag)
    
    With mTYPEICON
        .cbSize = Len(mTYPEICON)
        .picType = vbPicTypeIcon
        .hIcon = SHinfo.hIcon
    End With
    
    With mCLSID
        .id(8) = &HC0
        .id(15) = &H46
    End With
    
    Call OleCreatePictureIndirect(mTYPEICON, mCLSID, 1, GetFileIcon)

End Function


'------------------------------------------------------------------------------------------
' Procedure  : IsThere
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To check if the Picture is loaded
'------------------------------------------------------------------------------------------

Private Function IsThere(mPicture As StdPicture) As Boolean
On Error GoTo handle
    If Not mPicture.handle = 0 And Not mPicture.Height = 0 And Not mPicture.Width = 0 Then
        IsThere = True
    Else
        IsThere = False
    End If
Exit Function
handle:
    IsThere = False
End Function


'------------------------------------------------------------------------------------------------------------------------------------------
' Procedure : SplitToLines
' Auther    : Jim Jose
' Input     : Object, Text to split an parameters
' OutPut    : Splitted text array
' Purpose   : Split a string into lines by length!
'------------------------------------------------------------------------------------------------------------------------------------------

Private Function SplitToLines(ByVal sText As String, ByVal lLength As Long, _
                            Optional ByVal bFilterLines As Boolean = True) As String()
 Dim mArray() As String
 Dim mChar As String
 Dim mLine As String
 Dim lnCount As Long
 Dim xMax As String
 Dim mPos As Long
 Dim X As Long
 Dim lDone As Long

    If bFilterLines Then sText = Replace(sText, vbNewLine, vbNullString)
    xMax = Len(sText)
    
    For X = 1 To xMax
    
        mChar = Mid(sText, X, 1)

        If IsDelim(mChar) Then mPos = X - (lDone + 1)
        If TextWidth(mLine & mChar) >= lLength Or X = xMax Then
            If mPos = 0 Then mPos = X - (lDone + 1)
            ReDim Preserve mArray(lnCount)
            mArray(lnCount) = RTrim(LTrim(Mid(mLine, 1, mPos)))
            mLine = Mid(mLine, mPos + 1, Len(mLine) - mPos)
            lDone = lDone + mPos: mPos = 0
            lnCount = lnCount + 1
        End If
        
        mLine = mLine & mChar
        
    Next X

    mArray(lnCount - 1) = mArray(lnCount - 1) & mChar
    SplitToLines = mArray
    
End Function


'------------------------------------------------------------------------------------------------------------------------------------------
' Procedure : IsDelim
' Auther    : Rde
' Input     : Char
' OutPut    : IsDelim?
' Purpose   : Check if the input char is a Delimiter or not!
'------------------------------------------------------------------------------------------------------------------------------------------

Private Function IsDelim(Char As String) As Boolean
    Select Case Asc(Char) ' Upper/Lowercase letters,Underscore Not delimiters
    Case 65 To 90, 95, 97 To 122
        IsDelim = False
    Case Else: IsDelim = True ' Another Character Is delimiter
    End Select
End Function


'------------------------------------------------------------------------------------------------------------------------------------------
' Procedure : IsNT
' Auther    : Dana Seaman
' Input     : None
' OutPut    : NT?
' Purpose   : Check for the NT Platform
'------------------------------------------------------------------------------------------------------------------------------------------

Private Function IsNT() As Boolean
   Static m_bInit As Boolean
   Dim udtVer           As OSVERSIONINFO
   
   On Error Resume Next
   'Cache m_bIsNT on first execution
   If Not m_bInit Then
      m_bInit = True
      udtVer.dwOSVersionInfoSize = Len(udtVer)
      If GetVersionEx(udtVer) Then
         If udtVer.dwPlatformId = VER_PLATFORM_WIN32_NT Then
            m_bIsNT = True
         End If
      End If
   End If
   IsNT = m_bIsNT
   
End Function

'------------------------------------------------------------------------------------------
' Procedure  : SortSrtingArray
' Auther     : Jim Jose
' Input      : String Array + Order
' OutPut     : None
' Purpose    : To sort the String array Ascending/Descending
'------------------------------------------------------------------------------------------

Private Function SortSrtingArray(ByRef sArray() As String, _
                        ByVal vSortOrder As SortOrderEnum)
Dim X As Long
Dim Y As Long
Dim xMax As Long
Dim xStart As Long
Dim tmpStr As String
    
    On Error GoTo handle
    xStart = LBound(sArray)
    xMax = UBound(sArray)

    For X = xStart To xMax
        For Y = X + 1 To xMax
            If StrComp(sArray(X), sArray(Y), vbTextCompare) = vSortOrder Then
                tmpStr = sArray(X)
                sArray(X) = sArray(Y)
                sArray(Y) = tmpStr
            End If
        Next Y
    Next X

handle:
End Function


'------------------------------------------------------------------------------------------
' Procedure  : SortCollection
' Auther     : Jim Jose
' Input      : Collection to Sort + Order
' OutPut     : None
' Purpose    : To sort the Data-Collection Ascending/Descending
'------------------------------------------------------------------------------------------

Private Sub SortCollection(ByRef vCollection As Collection, _
                        ByVal vSortOrder As SortOrderEnum)
Dim X As Long
Dim vPos As Long
Dim vRtn  As Long
Dim vCount As Long
Dim vStart As Long
Dim vNewCount As Long
Dim vNew As New Collection
    
    ' Check Sort?
    If Not m_Mode = Mode_ListBox Then Exit Sub
    If vSortOrder = Sort_None Then Exit Sub
    
    ' Get current Count
    vStart = 1
    vCount = vCollection.Count

    ' Loop through Current collection
    For X = vStart To vCount
        
        ' Get new collection count
        vNewCount = vNew.Count
        
        ' Loop through new collection
        For vPos = 1 To vNewCount
        
            ' Compair each item in new collection
            vRtn = StrComp(vCollection(X), vNew(vPos), vbTextCompare)
            ' Escape with purpose
            If vRtn = vSortOrder Then Exit For
        
        Next vPos
        
        If X = vStart Or vPos = vNewCount + 1 Then
            ' New item at last
            vNew.add vCollection(X), "K " & X
        Else
            ' New item somewhere b/w
            vNew.add vCollection(X), "K " & X, vPos
        End If
        
    Next X
    
    ' Return Sorted Collection
    Set vCollection = vNew
    
End Sub

' -------------------------------------------------------------------------------------
' Procedure : BlendColor
' Type      : Property
' DateTime  : 03/02/2005
' Author    : Gary Noble
' Purpose   : Blends Two Colours Together
' Returns   : Long
' -------------------------------------------------------------------------------------

Private Property Get BlendColor(ByVal oColorFrom As OLE_COLOR, _
                               ByVal oColorTo As OLE_COLOR, _
                               Optional ByVal Alpha As Long = 128) As Long

Dim lSrcR As Long
Dim lSrcG As Long
Dim lSrcB As Long
Dim lDstR As Long
Dim lDstG As Long
Dim lDstB As Long
Dim lCFrom As Long
Dim lCTo As Long

    lCFrom = TranslateColor(oColorFrom)
    lCTo = TranslateColor(oColorTo)
    lSrcR = lCFrom And &HFF
    lSrcG = (lCFrom And &HFF00&) \ &H100&
    lSrcB = (lCFrom And &HFF0000) \ &H10000
    lDstR = lCTo And &HFF
    lDstG = (lCTo And &HFF00&) \ &H100&
    lDstB = (lCTo And &HFF0000) \ &H10000
    BlendColor = RGB(((lSrcR * Alpha) / 255) + ((lDstR * (255 - Alpha)) / 255), ((lSrcG * Alpha) / 255) + ((lDstG * (255 - Alpha)) / 255), ((lSrcB * Alpha) / 255) + ((lDstB * (255 - Alpha)) / 255))

End Property

' -------------------------------------------------------------------------------------
' Procedure : TranslateColor
' Type      : Function
' DateTime  : 03/02/2005
' Author    : Gary Noble
' Purpose   : Convert Automation color to Windows color
' Returns   : Long
' -------------------------------------------------------------------------------------

Private Function TranslateColor(ByVal oClr As OLE_COLOR, _
                               Optional hPal As Long = 0) As Long

    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = -1
    End If

End Function


'------------------------------------------------------------------------------------------------------------------------------------------
' Procedure : FillGradient
' Auther    : Jim Jose
' Input     : Hdc + Parameters
' OutPut    : None
' Purpose   : Middleout Gradients with Carls's DIB solution
'------------------------------------------------------------------------------------------------------------------------------------------

Private Sub FillGradient(ByVal hDC As Long, _
                         ByVal X As Long, _
                         ByVal Y As Long, _
                         ByVal Width As Long, _
                         ByVal Height As Long, _
                         ByVal Col1 As Long, _
                         ByVal Col2 As Long, _
                         ByVal GradientDirection As ListGradientDirectionEnum, _
                         Optional Right2Left As Boolean = True)
                         
Dim tmpCol  As Long
  
    ' Exit if needed
    If GradientDirection = Fill_None Then Exit Sub
    
    ' Right-To-Left
    If Right2Left Then
        tmpCol = Col1
        Col1 = Col2
        Col2 = tmpCol
    End If
    
    ' Translate system colors
    If Col1 < 0 Then Col1 = TranslateColor(Col1)
    If Col2 < 0 Then Col2 = TranslateColor(Col2)
    
    Select Case GradientDirection
        Case Fill_HorizontalMiddleOut
            DIBGradient hDC, X, Y, Width / 2, Height, Col1, Col2, Fill_Horizontal
            DIBGradient hDC, X + Width / 2 - 1, Y, Width / 2, Height, Col2, Col1, Fill_Horizontal

        Case Fill_VerticalMiddleOut
            DIBGradient hDC, X, Y, Width, Height / 2, Col1, Col2, Fill_Vertical
            DIBGradient hDC, X, Y + Height / 2 - 1, Width, Height / 2, Col2, Col1, Fill_Vertical

        Case Else
            DIBGradient hDC, X, Y, Width, Height, Col1, Col2, GradientDirection
    End Select
    
End Sub

'------------------------------------------------------------------------------------------------------------------------------------------
' Procedure : DIBGradient
' Auther    : Carls P.V.
' Input     : Hdc + Parameters
' OutPut    : None
' Purpose   : DIB solution for fast gradients
'------------------------------------------------------------------------------------------------------------------------------------------

Private Sub DIBGradient(ByVal hDC As Long, _
                         ByVal X As Long, _
                         ByVal Y As Long, _
                         ByVal Width As Long, _
                         ByVal Height As Long, _
                         ByVal Col1 As Long, _
                         ByVal Col2 As Long, _
                         ByVal GradientDirection As ListGradientDirectionEnum)

  Dim uBIH    As BITMAPINFOHEADER
  Dim lBits() As Long
  Dim lGrad() As Long
  
  Dim R1      As Long
  Dim G1      As Long
  Dim B1      As Long
  Dim R2      As Long
  Dim G2      As Long
  Dim B2      As Long
  Dim dR      As Long
  Dim dG      As Long
  Dim dB      As Long
  
  Dim Scan    As Long
  Dim i       As Long
  Dim iEnd    As Long
  Dim iOffset As Long
  Dim j       As Long
  Dim jEnd    As Long
  Dim iGrad   As Long
  
    '-- A minor check
    If (Width < 1 Or Height < 1) Then Exit Sub
    
    '-- Decompose Cols
    Col1 = Col1 And &HFFFFFF
    R1 = Col1 Mod &H100&
    Col1 = Col1 \ &H100&
    G1 = Col1 Mod &H100&
    Col1 = Col1 \ &H100&
    B1 = Col1 Mod &H100&
    Col2 = Col2 And &HFFFFFF
    R2 = Col2 Mod &H100&
    Col2 = Col2 \ &H100&
    G2 = Col2 Mod &H100&
    Col2 = Col2 \ &H100&
    B2 = Col2 Mod &H100&
    
    '-- Get Col distances
    dR = R2 - R1
    dG = G2 - G1
    dB = B2 - B1
    
    '-- Size gradient-Cols array
    Select Case GradientDirection
        Case [Fill_Horizontal]
            ReDim lGrad(0 To Width - 1)
        Case [Fill_Vertical]
            ReDim lGrad(0 To Height - 1)
        Case Else
            ReDim lGrad(0 To Width + Height - 2)
    End Select
    
    '-- Calculate gradient-Cols
    iEnd = UBound(lGrad())
    If (iEnd = 0) Then
        '-- Special case (1-pixel wide gradient)
        lGrad(0) = (B1 \ 2 + B2 \ 2) + 256 * (G1 \ 2 + G2 \ 2) + 65536 * (R1 \ 2 + R2 \ 2)
      Else
        For i = 0 To iEnd
            lGrad(i) = B1 + (dB * i) \ iEnd + 256 * (G1 + (dG * i) \ iEnd) + 65536 * (R1 + (dR * i) \ iEnd)
        Next i
    End If
    
    '-- Size DIB array
    ReDim lBits(Width * Height - 1) As Long
    iEnd = Width - 1
    jEnd = Height - 1
    Scan = Width
    
    '-- Render gradient DIB
    Select Case GradientDirection
        
        Case [Fill_Horizontal]
        
            For j = 0 To jEnd
                For i = iOffset To iEnd + iOffset
                    lBits(i) = lGrad(i - iOffset)
                Next i
                iOffset = iOffset + Scan
            Next j
        
        Case [Fill_Vertical]
        
            For j = jEnd To 0 Step -1
                For i = iOffset To iEnd + iOffset
                    lBits(i) = lGrad(j)
                Next i
                iOffset = iOffset + Scan
            Next j
            
        Case [Fill_DownwardDiagonal]
            
            iOffset = jEnd * Scan
            For j = 1 To jEnd + 1
                For i = iOffset To iEnd + iOffset
                    lBits(i) = lGrad(iGrad)
                    iGrad = iGrad + 1
                Next i
                iOffset = iOffset - Scan
                iGrad = j
            Next j
            
        Case [Fill_UpwardDiagonal]
            
            iOffset = 0
            For j = 1 To jEnd + 1
                For i = iOffset To iEnd + iOffset
                    lBits(i) = lGrad(iGrad)
                    iGrad = iGrad + 1
                Next i
                iOffset = iOffset + Scan
                iGrad = j
            Next j
    End Select
    
    '-- Define DIB header
    With uBIH
        .biSize = 40
        .biPlanes = 1
        .biBitCount = 32
        .biWidth = Width
        .biHeight = Height
    End With
    
    '-- Paint it!
    Call StretchDIBits(hDC, X, Y, Width, Height, 0, 0, Width, Height, lBits(0), uBIH, DIB_RGB_ColS, vbSrcCopy)

End Sub

'-----------------------------------------------------------------------------------------------------------
' Following is the API ScrollBar code by 'Gary Noble'. All the properties
' and functions are handles here.
'-----------------------------------------------------------------------------------------------------------
' Auther    :   Gary Noble
' Purpose   :   API ScrollBars
' About     :   This code uses Paul Carton's self subclaser. Create the ScrollBar(horz/vert),
'               on the event 'ReadProperties'. This code also enable 'MouseWheel'.
'               Gary, thanks a lot for providing this best code for API scrollBars...
'-----------------------------------------------------------------------------------------------------------

Friend Property Get SBVisible(ByVal eBar As EFSScrollBarConstants) As Boolean
    If (eBar = efsHorizontal) Then
        SBVisible = m_bVisibleHorz
    Else
        SBVisible = m_bVisibleVert
    End If
End Property

Friend Property Let SBVisible(ByVal eBar As EFSScrollBarConstants, ByVal bState As Boolean)
    If (eBar = efsHorizontal) Then
        m_bVisibleHorz = bState
    Else
        m_bVisibleVert = bState
    End If
    If (m_bNoFlatScrollBars) Then
        ShowScrollBar m_hWnd, eBar, Abs(bState)
    Else
        FlatSB_ShowScrollBar m_hWnd, eBar, Abs(bState)
    End If
End Property

Friend Property Get Orientation() As ScrollBarOrienationEnum
    SBOrientation = m_eOrientation
End Property

Friend Property Let SBOrientation(ByVal eOrientation As ScrollBarOrienationEnum)
    m_eOrientation = eOrientation
    pSBSetOrientation
End Property

Private Sub pSBSetOrientation()
    ShowScrollBar m_hWnd, SB_HORZ, Abs((m_eOrientation = Scroll_Both) Or (m_eOrientation = Scroll_Horizontal))
    ShowScrollBar m_hWnd, SB_VERT, Abs((m_eOrientation = Scroll_Both) Or (m_eOrientation = Scroll_Vertical))
End Sub

Private Sub SBRefresh()
    EnableScrollBar m_hWnd, SB_VERT, ESB_ENABLE_BOTH
End Sub

Private Sub pSBGetSI(ByVal eBar As EFSScrollBarConstants, ByRef tSI As SCROLLINFO, ByVal fMask As Long)
    Dim Lo As Long

    Lo = eBar
    tSI.fMask = fMask
    tSI.cbSize = LenB(tSI)

    If (m_bNoFlatScrollBars) Then
        GetScrollInfo m_hWnd, Lo, tSI
    Else
        FlatSB_GetScrollInfo m_hWnd, Lo, tSI
    End If

End Sub

Private Sub pSBLetSI(ByVal eBar As EFSScrollBarConstants, ByRef tSI As SCROLLINFO, ByVal fMask As Long)
    Dim Lo As Long

    Lo = eBar
    tSI.fMask = fMask
    tSI.cbSize = LenB(tSI)

    If (m_bNoFlatScrollBars) Then
        SetScrollInfo m_hWnd, Lo, tSI, True
    Else
        FlatSB_SetScrollInfo m_hWnd, Lo, tSI, True
    End If

End Sub

Friend Property Get SBStyle() As ScrollBarStyleEnum
    SBStyle = m_eStyle
End Property

Friend Property Let SBStyle(ByVal eStyle As ScrollBarStyleEnum)
    Dim lR As Long
    If (m_bNoFlatScrollBars) Then
        ' can't do it..
        'Debug.Print "Can't set non-regular style mode on this system - COMCTL32.DLL version < 4.71."
        Exit Property
    Else
        If (m_eOrientation = Scroll_Horizontal) Or (m_eOrientation = Scroll_Both) Then
            lR = FlatSB_SetScrollProp(m_hWnd, WSB_PROP_HSTYLE, eStyle, True)
        End If
        If (m_eOrientation = Scroll_Vertical) Or (m_eOrientation = Scroll_Both) Then
            lR = FlatSB_SetScrollProp(m_hWnd, WSB_PROP_VSTYLE, eStyle, True)
        End If
        'Debug.Print lR
        m_eStyle = eStyle
    End If

End Property

Friend Property Get SBSmallChange(ByVal eBar As EFSScrollBarConstants) As Long
    If (eBar = efsHorizontal) Then
        SBSmallChange = m_lSmallChangeHorz
    Else
        SBSmallChange = m_lSmallChangeVert
    End If
End Property

Friend Property Let SBSmallChange(ByVal eBar As EFSScrollBarConstants, ByVal lSmallChange As Long)
    If (eBar = efsHorizontal) Then
        m_lSmallChangeHorz = lSmallChange
    Else
        m_lSmallChangeVert = lSmallChange
    End If
End Property

Friend Property Get SBEnabled(ByVal eBar As EFSScrollBarConstants) As Boolean
    If (eBar = efsHorizontal) Then
        SBEnabled = m_bEnabledHorz
    Else
        SBEnabled = m_bEnabledVert
    End If
End Property

Friend Property Let SBEnabled(ByVal eBar As EFSScrollBarConstants, ByVal bEnabled As Boolean)
    Dim Lo As Long
    Dim lF As Long

    Lo = eBar
    If (bEnabled) Then
        lF = ESB_ENABLE_BOTH
    Else
        lF = ESB_DISABLE_BOTH
    End If
    If (m_bNoFlatScrollBars) Then
        EnableScrollBar m_hWnd, Lo, lF
    Else
        FlatSB_EnableScrollBar m_hWnd, Lo, lF
    End If

End Property

Friend Property Get SBMin(ByVal eBar As EFSScrollBarConstants) As Long
    Dim tSI As SCROLLINFO
    pSBGetSI eBar, tSI, SIF_RANGE
    SBMin = tSI.nMin
End Property

Friend Property Get SBMax(ByVal eBar As EFSScrollBarConstants) As Long
    Dim tSI As SCROLLINFO
    pSBGetSI eBar, tSI, SIF_RANGE Or SIF_PAGE
    SBMax = tSI.nMax                                  ' - tSI.nPage
End Property

Friend Property Get SBValue(ByVal eBar As EFSScrollBarConstants) As Long
    Dim tSI As SCROLLINFO
    pSBGetSI eBar, tSI, SIF_POS
    SBValue = tSI.nPos
End Property

Friend Property Get SBLargeChange(ByVal eBar As EFSScrollBarConstants) As Long
    Dim tSI As SCROLLINFO
    pSBGetSI eBar, tSI, SIF_PAGE
    SBLargeChange = tSI.nPage
End Property

Friend Property Let SBMin(ByVal eBar As EFSScrollBarConstants, ByVal iMin As Long)
    Dim tSI As SCROLLINFO
    tSI.nMin = iMin
    tSI.nMax = SBMax(eBar) + SBLargeChange(eBar)
    pSBLetSI eBar, tSI, SIF_RANGE
End Property

Friend Property Let SBMax(ByVal eBar As EFSScrollBarConstants, ByVal iMax As Long)
    Dim tSI As SCROLLINFO
    tSI.nMax = iMax + SBLargeChange(eBar)
    tSI.nMin = SBMin(eBar)
    pSBLetSI eBar, tSI, SIF_RANGE
End Property

Friend Property Let SBValue(ByVal eBar As EFSScrollBarConstants, ByVal iValue As Long)
    Dim tSI As SCROLLINFO
    If (iValue <> SBValue(eBar)) Then
        tSI.nPos = iValue
        pSBLetSI eBar, tSI, SIF_POS
        'ReDrawList
    End If
End Property

Friend Property Let SBLargeChange(ByVal eBar As EFSScrollBarConstants, ByVal iLargeChange As Long)
    Dim tSI As SCROLLINFO

    pSBGetSI eBar, tSI, SIF_ALL
    tSI.nMax = tSI.nMax - tSI.nPage + iLargeChange
    tSI.nPage = iLargeChange
    pSBLetSI eBar, tSI, SIF_PAGE Or SIF_RANGE
End Property

Friend Property Get SBCanBeFlat() As Boolean
    SBCanBeFlat = Not (m_bNoFlatScrollBars)
End Property

Private Sub pSBCreateScrollBar()
    Dim lR As Long
    Dim hParent As Long

    On Error Resume Next
    lR = InitialiseFlatSB(m_hWnd)
    If (Err.Number <> 0) Then
        'Can't find DLL entry point InitializeFlatSB in COMCTL32.DLL
        ' Means we have version prior to 4.71
        ' We get standard scroll bars.
        m_bNoFlatScrollBars = True
    Else
        SBStyle = m_eStyle
    End If
End Sub

Friend Sub SBCreate(ByVal hWndA As Long)
    pSBClearUp
    m_hWnd = hWndA
    pSBCreateScrollBar
End Sub

Private Sub pSBClearUp()
    If m_hWnd <> 0 Then
        On Error Resume Next
        ' Stop flat scroll bar if we have it:
        If Not (m_bNoFlatScrollBars) Then
            UninitializeFlatSB m_hWnd
        End If

        On Error GoTo 0
    End If
    m_hWnd = 0
    m_bInitialised = False
End Sub


'---------------------------------------------------------------------------------------------------------------------------------------------
' The following bytes are donated exclusively for Paul Caton's Subclassing
' We need this to track the movement information of the m_picCalendar and
' sizing/positioning of parent form
'---------------------------------------------------------------------------------------------------------------------------------------------
' Auther    : Paul Caton
' Purpose   : Advanced subclassing for UserControls (Self subclasser)
' Comment   : Thanks a Billion for this ever green piece of code on subclassing!!!
'---------------------------------------------------------------------------------------------------------------------------------------------

'Add a message to the table of those that will invoke a callback. You should Subclass_Start first and then add the messages
Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
'Parameters:
  'lng_hWnd  - The handle of the window for which the uMsg is to be added to the callback table
  'uMsg      - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
  With sc_aSubData(zIdx(lng_hWnd))
    If When And eMsgWhen.MSG_BEFORE Then
      Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If When And eMsgWhen.MSG_AFTER Then
      Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End If
  End With
End Sub

'Delete a message from the table of those that will invoke a callback.
Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
'Parameters:
  'lng_hWnd  - The handle of the window for which the uMsg is to be removed from the callback table
  'uMsg      - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback
  'When      - Whether the msg is to be removed from the before, after or both callback tables
  With sc_aSubData(zIdx(lng_hWnd))
    If When And eMsgWhen.MSG_BEFORE Then
      Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
    End If
    If When And eMsgWhen.MSG_AFTER Then
      Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
    End If
  End With
End Sub

'Return whether we're running in the IDE.
Private Function Subclass_InIDE() As Boolean
  'Debug.Assert zSetTrue(Subclass_InIDE)
End Function

'Start subclassing the passed window handle
Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
'Parameters:
  'lng_hWnd  - The handle of the window to be subclassed
'Returns;
  'The sc_aSubData() index
  Const CODE_LEN              As Long = 200                                             'Length of the machine code in bytes
  Const FUNC_CWP              As String = "CallWindowProcA"                             'We use CallWindowProc to call the original WndProc
  Const FUNC_EBM              As String = "EbMode"                                      'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
  Const FUNC_SWL              As String = "SetWindowLongA"                              'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
  Const MOD_USER              As String = "user32"                                      'Location of the SetWindowLongA & CallWindowProc functions
  Const MOD_VBA5              As String = "vba5"                                        'Location of the EbMode function if running VB5
  Const MOD_VBA6              As String = "vba6"                                        'Location of the EbMode function if running VB6
  Const PATCH_01              As Long = 18                                              'Code buffer offset to the location of the relative address to EbMode
  Const PATCH_02              As Long = 68                                              'Address of the previous WndProc
  Const PATCH_03              As Long = 78                                              'Relative address of SetWindowsLong
  Const PATCH_06              As Long = 116                                             'Address of the previous WndProc
  Const PATCH_07              As Long = 121                                             'Relative address of CallWindowProc
  Const PATCH_0A              As Long = 186                                             'Address of the owner object
  Static aBuf(1 To CODE_LEN)  As Byte                                                   'Static code buffer byte array
  Static pCWP                 As Long                                                   'Address of the CallWindowsProc
  Static pEbMode              As Long                                                   'Address of the EbMode IDE break/stop/running function
  Static pSWL                 As Long                                                   'Address of the SetWindowsLong function
  Dim i                       As Long                                                   'Loop index
  Dim j                       As Long                                                   'Loop index
  Dim nSubIdx                 As Long                                                   'Subclass data index
  Dim sHex                    As String                                                 'Hex code string
  
'If it's the first time through here..
  If aBuf(1) = 0 Then
  
'The hex pair machine code representation.
    sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & _
           "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & _
           "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & _
           "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"

'Convert the string from hex pairs to bytes and store in the static machine code buffer
    i = 1
    Do While j < CODE_LEN
      j = j + 1
      aBuf(j) = Val("&H" & Mid$(sHex, i, 2))                                            'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
      i = i + 2
    Loop                                                                                'Next pair of hex characters
    
'Get API function addresses
    If Subclass_InIDE Then                                                              'If we're running in the VB IDE
      aBuf(16) = &H90                                                                   'Patch the code buffer to enable the IDE state code
      aBuf(17) = &H90                                                                   'Patch the code buffer to enable the IDE state code
      pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                                           'Get the address of EbMode in vba6.dll
      If pEbMode = 0 Then                                                               'Found?
        pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                                         'VB5 perhaps
      End If
    End If
    
    pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                                                'Get the address of the CallWindowsProc function
    pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                                                'Get the address of the SetWindowLongA function
    ReDim sc_aSubData(0 To 0) As tSubData                                               'Create the first sc_aSubData element
  Else
    nSubIdx = zIdx(lng_hWnd, True)
    If nSubIdx = -1 Then                                                                'If an sc_aSubData element isn't being re-cycled
      nSubIdx = UBound(sc_aSubData()) + 1                                               'Calculate the next element
      ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData                              'Create a new sc_aSubData element
    End If
    
    Subclass_Start = nSubIdx
  End If

  With sc_aSubData(nSubIdx)
    .hwnd = lng_hWnd                                                                    'Store the hWnd
    .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)                                       'Allocate memory for the machine code WndProc
    .nAddrOrig = SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrSub)                          'Set our WndProc in place
    Call RtlMoveMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)                              'Copy the machine code from the static byte array to the code array in sc_aSubData
    Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)                                        'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
    Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)                                     'Original WndProc address for CallWindowProc, call the original WndProc
    Call zPatchRel(.nAddrSub, PATCH_03, pSWL)                                           'Patch the relative address of the SetWindowLongA api function
    Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)                                     'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
    Call zPatchRel(.nAddrSub, PATCH_07, pCWP)                                           'Patch the relative address of the CallWindowProc api function
    Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))                                     'Patch the address of this object instance into the static machine code buffer
  End With
End Function

'Stop all subclassing
Private Sub Subclass_StopAll()
  Dim i As Long
  
  i = UBound(sc_aSubData())                                                             'Get the upper bound of the subclass data array
  Do While i >= 0                                                                       'Iterate through each element
    With sc_aSubData(i)
      If .hwnd <> 0 Then                                                                'If not previously Subclass_Stop'd
        Call Subclass_Stop(.hwnd)                                                       'Subclass_Stop
      End If
    End With
    
    i = i - 1                                                                           'Next element
  Loop
End Sub

'Stop subclassing the passed window handle
Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
'Parameters:
  'lng_hWnd  - The handle of the window to stop being subclassed
  With sc_aSubData(zIdx(lng_hWnd))
    Call SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrOrig)                                 'Restore the original WndProc
    Call zPatchVal(.nAddrSub, PATCH_05, 0)                                              'Patch the Table B entry count to ensure no further 'before' callbacks
    Call zPatchVal(.nAddrSub, PATCH_09, 0)                                              'Patch the Table A entry count to ensure no further 'after' callbacks
    Call GlobalFree(.nAddrSub)                                                          'Release the machine code memory
    .hwnd = 0                                                                           'Mark the sc_aSubData element as available for re-use
    .nMsgCntB = 0                                                                       'Clear the before table
    .nMsgCntA = 0                                                                       'Clear the after table
    Erase .aMsgTblB                                                                     'Erase the before table
    Erase .aMsgTblA                                                                     'Erase the after table
  End With
End Sub

'Worker sub for Subclass_AddMsg
Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
  Dim nEntry  As Long                                                                   'Message table entry index
  Dim nOff1   As Long                                                                   'Machine code buffer offset 1
  Dim nOff2   As Long                                                                   'Machine code buffer offset 2
  
  If uMsg = ALL_MESSAGES Then                                                           'If all messages
    nMsgCnt = ALL_MESSAGES                                                              'Indicates that all messages will callback
  Else                                                                                  'Else a specific message number
    Do While nEntry < nMsgCnt                                                           'For each existing entry. NB will skip if nMsgCnt = 0
      nEntry = nEntry + 1
      
      If aMsgTbl(nEntry) = 0 Then                                                       'This msg table slot is a deleted entry
        aMsgTbl(nEntry) = uMsg                                                          'Re-use this entry
        Exit Sub                                                                        'Bail
      ElseIf aMsgTbl(nEntry) = uMsg Then                                                'The msg is already in the table!
        Exit Sub                                                                        'Bail
      End If
    Loop                                                                                'Next entry

    nMsgCnt = nMsgCnt + 1                                                               'New slot required, bump the table entry count
    ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                                        'Bump the size of the table.
    aMsgTbl(nMsgCnt) = uMsg                                                             'Store the message number in the table
  End If

  If When = eMsgWhen.MSG_BEFORE Then                                                    'If before
    nOff1 = PATCH_04                                                                    'Offset to the Before table
    nOff2 = PATCH_05                                                                    'Offset to the Before table entry count
  Else                                                                                  'Else after
    nOff1 = PATCH_08                                                                    'Offset to the After table
    nOff2 = PATCH_09                                                                    'Offset to the After table entry count
  End If

  If uMsg <> ALL_MESSAGES Then
    Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))                                    'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
  End If
  Call zPatchVal(nAddr, nOff2, nMsgCnt)                                                 'Patch the appropriate table entry count
End Sub

'Return the memory address of the passed function in the passed dll
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
  zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
  'Debug.Assert zAddrFunc                                                                'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
End Function

'Worker sub for Subclass_DelMsg
Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
  Dim nEntry As Long
  
  If uMsg = ALL_MESSAGES Then                                                           'If deleting all messages
    nMsgCnt = 0                                                                         'Message count is now zero
    If When = eMsgWhen.MSG_BEFORE Then                                                  'If before
      nEntry = PATCH_05                                                                 'Patch the before table message count location
    Else                                                                                'Else after
      nEntry = PATCH_09                                                                 'Patch the after table message count location
    End If
    Call zPatchVal(nAddr, nEntry, 0)                                                    'Patch the table message count to zero
  Else                                                                                  'Else deleteting a specific message
    Do While nEntry < nMsgCnt                                                           'For each table entry
      nEntry = nEntry + 1
      If aMsgTbl(nEntry) = uMsg Then                                                    'If this entry is the message we wish to delete
        aMsgTbl(nEntry) = 0                                                             'Mark the table slot as available
        Exit Do                                                                         'Bail
      End If
    Loop                                                                                'Next entry
  End If
End Sub

'Get the sc_aSubData() array index of the passed hWnd
Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
'Get the upper bound of sc_aSubData() - If you get an error here, you're probably Subclass_AddMsg-ing before Subclass_Start
  zIdx = UBound(sc_aSubData)
  Do While zIdx >= 0                                                                    'Iterate through the existing sc_aSubData() elements
    With sc_aSubData(zIdx)
      If .hwnd = lng_hWnd Then                                                          'If the hWnd of this element is the one we're looking for
        If Not bAdd Then                                                                'If we're searching not adding
          Exit Function                                                                 'Found
        End If
      ElseIf .hwnd = 0 Then                                                             'If this an element marked for reuse.
        If bAdd Then                                                                    'If we're adding
          Exit Function                                                                 'Re-use it
        End If
      End If
    End With
    zIdx = zIdx - 1                                                                     'Decrement the index
  Loop
  
  If Not bAdd Then
    'Debug.Assert False                                                                  'hWnd not found, programmer error
  End If

'If we exit here, we're returning -1, no freed elements were found
End Function

'Patch the machine code buffer at the indicated offset with the relative address to the target address.
Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
  Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

'Patch the machine code buffer at the indicated offset with the passed value
Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
  Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

'Worker function for Subclass_InIDE
Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
  zSetTrue = True
  bValue = True
End Function

'Return the upper 16 bits of the passed 32 bit value
Private Function WordHi(lngValue As Long) As Long
  If (lngValue And &H80000000) = &H80000000 Then
    WordHi = ((lngValue And &H7FFF0000) \ &H10000) Or &H8000&
  Else
    WordHi = (lngValue And &HFFFF0000) \ &H10000
  End If
End Function

'Return the lower 16 bits of the passed 32 bit value
Private Function WordLo(lngValue As Long) As Long
  WordLo = (lngValue And &HFFFF&)
End Function

'Determine if the passed function is supported
Private Function IsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean
  Dim hMod        As Long
  Dim bLibLoaded  As Boolean

  hMod = GetModuleHandleA(sModule)

  If hMod = 0 Then
    hMod = LoadLibraryA(sModule)
    If hMod Then
      bLibLoaded = True
    End If
  End If

  If hMod Then
    If GetProcAddress(hMod, sFunction) Then
      IsFunctionExported = True
    End If
  End If

  If bLibLoaded Then
    Call FreeLibrary(hMod)
  End If
End Function

'Track the mouse leaving the indicated window
Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)
  Dim tme As TRACKMOUSEEVENT_STRUCT
  
  If bTrack Then
    With tme
      .cbSize = Len(tme)
      .dwFlags = TME_LEAVE
      .hwndTrack = lng_hWnd
    End With

    If bTrackUser32 Then
      Call TrackMouseEvent(tme)
    Else
      Call TrackMouseEventComCtl(tme)
    End If
  End If
  
End Sub


Private Function GetVolumeLabel(ByVal sDrive As String) As String

    Dim sBuffer As String
    Dim sSysName As String
    Dim lResult As Long
    Dim lSysFlags As Long
    Dim lComponentLength As Long
    Dim mSerial As Long
    Dim lRtn As Long
    
    lRtn = GetDriveType(sDrive)
    
    Select Case lRtn
        Case DRIVE_REMOVABLE
            GetVolumeLabel = "Floppy"
        Case DRIVE_CDROM, DRIVE_FIXED, DRIVE_RAMDISK, DRIVE_REMOTE
            sBuffer = String$(256, 0)
            sSysName = String$(256, 0)
            lResult = GetVolumeInformation(sDrive, sBuffer, 255, mSerial, lComponentLength, lSysFlags, sSysName, 255)
        
            If Not lResult = 0 Then
                ' retrieve the information
                sBuffer = Left$(sBuffer, InStr(sBuffer, Chr$(0)) - 1)
                If sBuffer = vbNullString Then
                    GetVolumeLabel = "Local Disc"
                Else
                    GetVolumeLabel = StrConv(sBuffer, vbProperCase)
                End If
            Else
                If lRtn = DRIVE_CDROM Then GetVolumeLabel = "CD Drive"
            End If
            
    End Select
    
End Function
