Attribute VB_Name = "mUniMenu"
Option Explicit

Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoW" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Long, lpMenuItemInfo As MENUITEMINFOW) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoW" (ByVal hMenu As Long, ByVal un As Long, ByVal BOOL As Boolean, lpcMenuItemInfo As MENUITEMINFOW) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Private Declare Function SetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPos As Long) As Long

Private Type MENUITEMINFOW
   cbSize           As Long
   fMask            As Long
   fType            As Long
   fState           As Long
   wID              As Long
   hSubMenu         As Long
   hbmpChecked      As Long
   hbmpUnchecked    As Long
   dwItemData       As Long
   dwTypeData       As Long
   cch              As Long
   hbmpItem         As Long
End Type

Private Const MIIM_TYPE = &H10
Private Const MIIM_DATA = &H20
Private Const MF_UNCHECKED = &H0&
Private Const MF_CHECKED = &H8&
Private Const MF_BYPOSITION = &H400&

Public Sub InitMenu(hwnd As Long)
    Dim hMenu&
    hMenu = GetMenu(hwnd)   'lay handle cua thanh menu trong cua so ung dung
    VietnameseMenu hMenu    'lay tung menu con trong thanh menu cua ung dung
End Sub

Private Sub VietnameseMenu(ByVal hMenu As Long)
    Dim hSubMenu&, i%, nCnt%, sTmp$, sStr$
    Dim MII As MENUITEMINFOW
    
    
    sStr = String(&HFF, 0)
    nCnt = GetMenuItemCount(hMenu)  'dem so menu con trong thanh menu cua cua so ung dung
    If nCnt Then
        For i = 0 To nCnt - 1
            MII.cbSize = LenB(MII)
            MII.fMask = MIIM_TYPE Or MIIM_DATA
            MII.dwTypeData = StrPtr(sStr)  ' String(&HFF, 0)
            MII.cch = Len(sStr)  'MII.dwTypeData)
            MII.hbmpChecked = MF_CHECKED Or MF_UNCHECKED
            
'lay caption cua Menu
            GetMenuItemInfo hMenu, i, True, MII
            sTmp = Left$(sStr, MII.cch)  ' MII.dwTypeData, MII.cch)
            
            If sTmp <> "" Then
                sTmp = zToUnicode(sTmp)
                SetUniMenu sTmp, hMenu, i   'thuc hien gan caption cho menu vua tim duoc
            End If
            
'lay Menu con cua mot MenuItem
            hSubMenu = GetSubMenu(hMenu, i)     'lay handle menu con cua menu hien dang xu ly
            If hSubMenu Then    'neu tim thay handle thi goi de quy de xu ly caption cac menu item con trong menu cha dang xu ly
                VietnameseMenu hSubMenu
            End If
        Next
    End If
End Sub

Public Sub SetUniMenu(sCaption As String, MnuHwnd As Long, ByVal mnuItem As Long, Optional ByVal mnuParentItem As Long = -1, Optional isDefault As Boolean = False)
    Dim hMenu As Long
    Dim mInfo As MENUITEMINFOW
    
    If isDefault Then SetMenuDefaultItem hMenu, mnuItem, 1
    With mInfo
        .cbSize = Len(mInfo)
        .fType = &H200
        .fMask = &H10
        .dwTypeData = StrPtr(sCaption)
    End With
    SetMenuItemInfo MnuHwnd, mnuItem, 1, mInfo
End Sub

Public Sub SetIconForMenu(HwndForm As Long, MenuNumber As Integer, SubMenuItemCount1 As Integer, Optional SubMenuItemCount2 As Integer, Optional SubMenuItemCount3 As Integer, Optional Icon As Picture, Optional isDefault As Boolean)
On Error GoTo Err
    Dim hMainMenu As Long, hSubMenu1 As Long, hSubMenu2 As Long, hSubMenu3 As Long
    
    MenuNumber = MenuNumber - 1
    SubMenuItemCount1 = SubMenuItemCount1 - 1
    SubMenuItemCount2 = SubMenuItemCount2 - 1
    SubMenuItemCount3 = SubMenuItemCount3 - 1
    
    hMainMenu = GetMenu(HwndForm)       'lay menu cua form
    
    
    If SubMenuItemCount1 >= 0 Then hSubMenu1 = GetSubMenu(hMainMenu, MenuNumber)        'lay menu con thu 1
    If SubMenuItemCount2 >= 0 Then hSubMenu2 = GetSubMenu(hSubMenu1, SubMenuItemCount1) 'lay menu con thu 2
    If SubMenuItemCount3 >= 0 Then hSubMenu3 = GetSubMenu(hSubMenu2, SubMenuItemCount2) 'lay menu con thu 3
    
'neu chon Icon cho mot Menu khong ton tai trong Menu hien tai thi thoat khoi thuc tuc
    If (hSubMenu3 = 0 And SubMenuItemCount3 >= 0) Or (hSubMenu2 = 0 And SubMenuItemCount2 >= 0) Or (hSubMenu1 = 0 And SubMenuItemCount1 >= 0) Then Exit Sub
    
'neu chon dat Icon cho menu con cap 3
    If hSubMenu3 <> 0 Then
        If isDefault Then SetMenuDefaultItem hSubMenu3, SubMenuItemCount3, 1
        SetMenuItemBitmaps hSubMenu3, SubMenuItemCount3, MF_BYPOSITION, Icon, Icon
        Exit Sub
    End If

'neu chon dat Icon cho menu con cap 2
    If hSubMenu2 <> 0 Then
        If isDefault Then SetMenuDefaultItem hSubMenu2, SubMenuItemCount2, 1
        SetMenuItemBitmaps hSubMenu2, SubMenuItemCount2, MF_BYPOSITION, Icon, Icon
        Exit Sub
    End If
    
'neu chon dat Icon cho menu con cap 1
    If hSubMenu1 <> 0 Then
        If isDefault Then SetMenuDefaultItem hSubMenu1, SubMenuItemCount1, 1
        SetMenuItemBitmaps hSubMenu1, SubMenuItemCount1, MF_BYPOSITION, Icon, Icon
        Exit Sub
    End If

Err:
'loi xay ra khi chon menu can dat Icon ma khong dat icon
End Sub

