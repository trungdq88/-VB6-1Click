VERSION 5.00
Begin VB.UserControl UniMenu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   420
   FillColor       =   &H8000000F&
   ForeColor       =   &H8000000F&
   InvisibleAtRuntime=   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   435
   ScaleWidth      =   420
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   360
      Left            =   30
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   0
      Top             =   40
      Width           =   360
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      Height          =   435
      Left            =   0
      Top             =   0
      Width           =   420
   End
End
Attribute VB_Name = "UniMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub InitUnicodeMenu(Optional HwndForm As Long = 0)
On Error Resume Next

    InitMenu IIf(HwndForm <> 0, HwndForm, UserControl.Parent.hwnd)
End Sub

Public Sub SetMenuIcon(HwndForm As Long, MenuNumber As Integer, SubMenuItemCount1 As Integer, Optional SubMenuItemCount2 As Integer, Optional SubMenuItemCount3 As Integer, Optional Icon As Picture, Optional isDefault As Boolean)
On Error Resume Next

    Call SetIconForMenu(HwndForm, MenuNumber, SubMenuItemCount1, SubMenuItemCount2, SubMenuItemCount3, Icon, isDefault)
End Sub

Private Sub UserControl_Resize()
    UserControl.Size 420, 435
End Sub

Public Sub About()
Attribute About.VB_UserMemId = -552
On Error Resume Next:   frmAbout.Show vbModal
End Sub
