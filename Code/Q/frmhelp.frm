VERSION 5.00
Object = "{E8FDD05C-3067-4198-8AEC-1A013A46ABDD}#1.0#0"; "FVUnicodeControl.ocx"
Begin VB.Form frmHelpMe 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help"
   ClientHeight    =   7785
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   6255
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmhelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FVUnicodeControl.FVistaUniCheckbox chkNotView 
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   7080
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Kho6ng Hie63n Thi5 Va2o La62n Sau."
      ForeColor       =   0
   End
   Begin FVUnicodeControl.FVistaUniButton cmdHelpChiTiet 
      Height          =   255
      Left            =   3960
      TabIndex        =   3
      Top             =   7440
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      BackColor       =   -2147483633
      ButtonStyle     =   3
      Caption         =   "Hu7o71ng Da64n Chi Tie61t"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin FVUnicodeControl.FVistaUniButton cmdExit 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   7440
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      BackColor       =   -2147483633
      ButtonStyle     =   3
      Caption         =   "D9o1ng"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin FVUnicodeControl.FVistaUniLabel lblHelp 
      Height          =   6375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   11245
      BorderStyle     =   1
      Caption         =   "D9a6y la2 no65i dung hu7o71ng da64n"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483635
   End
   Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   661
      Alignment       =   1
      Caption         =   "Co1 Gi2 Trong Phie6n Ba3n 1.2 ?"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711935
   End
End
Attribute VB_Name = "frmHelpMe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkNotView_Click()
If chkNotView.Value = True Then
    SaveString HKEY_CURRENT_USER, "Software\1Click", "HelpUpdate", "1"
Else
    DeleteValue HKEY_CURRENT_USER, "Software\1Click", "HelpUpdate"
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdHelpChiTiet_Click()
UniMsgBox ChrW$(&H4D) & ChrW$(&HE1) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&HED) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H1EA7) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H70) & ChrW$(&H68) & ChrW$(&H1EA3) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&H1EBF) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H1ED1) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H49) & ChrW$(&H6E) & ChrW$(&H74) & ChrW$(&H65) _
& ChrW$(&H72) & ChrW$(&H6E) & ChrW$(&H65) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1EC3) & ChrW$(&H20) & ChrW$(&H78) & ChrW$(&H65) & ChrW$(&H6D) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1B0) & ChrW$(&H1EE3) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H68) & ChrW$(&H1B0) & ChrW$(&H1EDB) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H1EAB) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H69) & ChrW$(&H1EBF) _
 & ChrW$(&H74) & ChrW$(&H2E) & vbCrLf _
 & ChrW$(&H42) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H169) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&HF3) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H68) & ChrW$(&H1EC3) & ChrW$(&H20) & ChrW$(&H78) & ChrW$(&H65) & ChrW$(&H6D) & ChrW$(&H20) & ChrW$(&H68) & ChrW$(&H1B0) & ChrW$(&H1EDB) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H1EAB) & ChrW$(&H6E) _
& ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H69) & ChrW$(&H1EBF) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H1EDF) & ChrW$(&H20) & ChrW$(&H46) & ChrW$(&H69) & ChrW$(&H6C) & ChrW$(&H65) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&HE9) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&HE8) & ChrW$(&H6D) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H68) & ChrW$(&H65) & ChrW$(&H6F) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H1B0) & ChrW$(&H1A1) & ChrW$(&H6E) _
 & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H72) & ChrW$(&HEC) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H2E), vbOKOnly, "Thông Báo", Me.hwnd
frmHelpChi.Show
Unload Me
End Sub

Private Sub Form_Load()
If GetString(HKEY_CURRENT_USER, "Software\1Click", "HelpUpdate") = "1" Then chkNotView.Value = True

lblHelp.Caption = "**********************************" & vbCrLf _
& "*** Ch" & ChrW(7913) & "c n" & ChrW(259) & "ng c" & ChrW(224) & "i " & ChrW(273) & ChrW(7863) & "t chung:" & vbCrLf _
& "       - " & ChrW(272) & ChrW(432) & "a ra th" & ChrW(244) & "ng tin chung v" & ChrW(7873) & " m" & ChrW(225) & "y t" & ChrW(237) & "nh c" & ChrW(7911) & "a b" & ChrW(7841) & "n." & vbCrLf _
& vbCrLf _
& "*** Ch" & ChrW(7913) & "c n" & ChrW(259) & "ng Ph" & ChrW(7909) & "c H" & ChrW(7891) & "i && S" & ChrW(7919) & "a Ch" & ChrW(7919) & "a:" & vbCrLf _
& "       - " & "N" & ChrW(226) & "ng C" & ChrW(7845) & "p + " & "T" & ChrW(259) & "ng t" & ChrW(7889) & "c ch" & ChrW(7913) & "c n" & ChrW(259) & "ng x" & ChrW(243) & "a Autorun." & vbCrLf _
& "       - " & "N" & ChrW(226) & "ng C" & ChrW(7845) & "p + " & " T" & ChrW(259) & "ng t" & ChrW(7889) & "c ch" & ChrW(7913) & "c n" & ChrW(259) & "ng ch" & ChrW(7889) & "ng Autorun." & vbCrLf _
& "       - Th" & ChrW(234) & "m ch" & ChrW(7913) & "c n" & ChrW(259) & "ng ph" & ChrW(7909) & "c h" & ChrW(7891) & "i Registry v" & ChrW(7873) & " t" & ChrW(236) & "nh tr" & ChrW(7841) & "ng t" & ChrW(7889) & "t nh" & ChrW(7845) & "t." & vbCrLf _
& "       - Th" & ChrW(234) & "m ch" & ChrW(7913) & "c n" & ChrW(259) & "ng Thay " & ChrW(273) & ChrW(7893) & "i Logo v" & ChrW(224) & " Th" & ChrW(244) & "ng Tin H" & ChrW(7879) & " Th" & ChrW(7889) & "ng." & vbCrLf _
& vbCrLf _
& "*** Ch" & ChrW(7913) & "c n" & ChrW(259) & "ng Di" & ChrW(7879) & "t Virus && B" & ChrW(7843) & "o M" & ChrW(7853) & "t:" & vbCrLf _
& "       - C" & ChrW(7853) & "p nh" & ChrW(7853) & "t th" & ChrW(234) & "m m" & ChrW(7897) & "t s" & ChrW(7889) & " lo" & ChrW(7841) & "i Virus m" & ChrW(7899) & "i. T" & ChrW(259) & "ng t" & ChrW(7889) & "c di" & ChrW(7879) & "t Virus nhanh g" & ChrW(7845) & "p 2 l" & ChrW(7847) & "n." & vbCrLf _
& vbCrLf _
& "*** Ch" & ChrW(7913) & "c n" & ChrW(259) & "ng Ki" & ChrW(7875) & "m Tra M" & ChrW(225) & "y T" & ChrW(237) & "nh:" & vbCrLf _
& "       - Ki" & ChrW(7875) & "m tra c" & ChrW(225) & "c th" & ChrW(244) & "ng s" & ChrW(7889) & " v" & ChrW(224) & " c" & ChrW(225) & "c File quan tr" & ChrW(7885) & "ng c" & ChrW(7911) & "a h" & ChrW(7879) & " th" & ChrW(244) & "ng " & ChrW(273) & ChrW(7875) & " " & ChrW(273) & ChrW(432) & "a ra File Log v" & ChrW(7873) & " t" & ChrW(236) & "nh tr" & ChrW(7841) & "ng m" & ChrW(225) & "y t" & ChrW(237) & "nh c" & ChrW(7911) & "a b" & ChrW(7841) & "n. File Log c" & ChrW(243) & " kh" & ChrW(7843) & " n" & ChrW(259) & "ng " & ChrW(273) & ChrW(225) & "nh gi" & ChrW(225) & " 70% t" & ChrW(236) & "nh tr" & ChrW(7841) & "ng hi" & ChrW(7879) & "n t" & ChrW(7841) & "i c" & ChrW(7911) & "a m" & ChrW(225) & "y t" & ChrW(237) & "nh." & vbCrLf _
& vbCrLf _
& "*** Ch" & ChrW(7913) & "c n" & ChrW(259) & "ng Qu" & ChrW(7843) & "n L" & ChrW(253) & " File:" & vbCrLf _
& "       - Qu" & ChrW(7843) & "n l" & ChrW(253) & " v" & ChrW(224) & " thao t" & ChrW(225) & "c v" & ChrW(7899) & "i File v" & ChrW(224) & " th" & ChrW(432) & " m" & ChrW(7909) & "c. Ch" & ChrW(7913) & "c n" & ChrW(259) & "ng n" & ChrW(224) & "y c" & ChrW(243) & " th" & ChrW(7875) & " thay th" & ChrW(7871) & " cho Explorer trong tr" & ChrW(432) & ChrW(7901) & "ng h" & ChrW(7907) & "p m" & ChrW(225) & "y t" & ChrW(237) & "nh b" & ChrW(7883) & " nhi" & ChrW(7877) & "m Virus, kh" & ChrW(244) & "ng xem " & ChrW(273) & ChrW(432) & ChrW(7907) & "c c" & ChrW(225) & "c File " & ChrW(7849) & "n." & vbclrf _
& vbCrLf _
& vbCrLf _
& "*** Ch" & ChrW(7913) & "c n" & ChrW(259) & "ng T" & ChrW(236) & "m Ki" & ChrW(7871) & "m:" & vbCrLf _
& "       - T" & ChrW(236) & "m ki" & ChrW(7871) & "m t" & ChrW(7889) & "c " & ChrW(273) & ChrW(7897) & " cao: Gi" & ChrW(250) & "p t" & ChrW(236) & "m ki" & ChrW(7871) & "m c" & ChrW(225) & "c lo" & ChrW(7841) & "i File v" & ChrW(7899) & "i t" & ChrW(7889) & "c " & ChrW(273) & ChrW(7897) & " nhanh." & vbCrLf _
& "       - T" & ChrW(236) & "m ki" & ChrW(7871) & "m c" & ChrW(225) & "c File c" & ChrW(243) & " t" & ChrW(234) & "n tr" & ChrW(249) & "ng v" & ChrW(7899) & "i t" & ChrW(234) & "n th" & ChrW(432) & " m" & ChrW(7909) & "c ch" & ChrW(7913) & "a n" & ChrW(243) & ": Nh" & ChrW(7919) & "ng file d" & ChrW(7841) & "ng n" & ChrW(224) & "y th" & ChrW(432) & ChrW(7901) & "ng l" & ChrW(224) & " Virus gi" & ChrW(7843) & " d" & ChrW(7841) & "ng, ch" & ChrW(432) & ChrW(417) & "ng tr" & ChrW(236) & "nh s" & ChrW(7869) & " t" & ChrW(236) & "m c" & ChrW(225) & "c file thu" & ChrW(7897) & "c lo" & ChrW(7841) & "i n" & ChrW(224) & "y v" & ChrW(224) & " cho ph" & ChrW(233) & "p x" & ChrW(243) & "a nhanh ch" & ChrW(243) & "ng." & vbCrLf _
& "       - T" & ChrW(236) & "m ki" & ChrW(7871) & "m File theo Icon: Cho ph" & ChrW(233) & "p t" & ChrW(236) & "m ki" & ChrW(7871) & "m nh" & ChrW(7919) & "ng file c" & ChrW(243) & " Icon gi" & ChrW(7889) & "ng (Ho" & ChrW(7863) & "c g" & ChrW(7847) & "n gi" & ChrW(7889) & "ng) so v" & ChrW(7899) & "i icon do ng" & ChrW(432) & ChrW(7901) & "i d" & ChrW(249) & "ng ch" & ChrW(7881) & " " & ChrW(273) & ChrW(7883) & "nh. Ch" & ChrW(7913) & "c n" & ChrW(259) & "ng n" & ChrW(224) & "y c" & ChrW(243) & " th" & ChrW(7875) & " gi" & ChrW(250) & "p qu" & ChrW(233) & "t c" & ChrW(225) & "c lo" & ChrW(7841) & "i virus gi" & ChrW(7843) & " d" & ChrW(7841) & "ng th" & ChrW(432) & " m" & ChrW(7909) & "c, gi" & ChrW(7843) & " d" & ChrW(7841) & "ng file h" & ChrW(7879) & " th" & ChrW(7889) & "ng,..."
End Sub

Private Sub FVistaUniCheckbox1_Click()

End Sub
