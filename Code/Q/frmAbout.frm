VERSION 5.00
Object = "{E8FDD05C-3067-4198-8AEC-1A013A46ABDD}#1.0#0"; "FVUnicodeControl.ocx"
Begin VB.Form frmAbout 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About 1Click."
   ClientHeight    =   3495
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FVUnicodeControl.FVistaUniButton cmdGotoHomePages 
      Height          =   375
      Left            =   600
      TabIndex        =   14
      Top             =   3000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BackColor       =   -2147483626
      ButtonStyle     =   3
      Caption         =   "Trang Chu3"
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
   Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel5 
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      Alignment       =   2
      Caption         =   "Ngo6n ngu74 su73 du5ng:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483635
   End
   Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel3 
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   840
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   450
      Alignment       =   1
      Caption         =   "Chu7o7ng tri2nh su73a chu74a, phu5c ho62i ca1c chu71c na8ng cu73a ma1y ti1nh"
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
   Begin FVUnicodeControl.FVistaUniButton cmdExit 
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   3000
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      BackColor       =   -2147483633
      ButtonStyle     =   3
      Caption         =   "D9o1ng"
      Effects         =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicOpacity      =   0
   End
   Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel6 
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      Caption         =   "D9inh Quang Trung"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel4 
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      Alignment       =   2
      Caption         =   "Ta1c Gia3:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483635
   End
   Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel2 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BackColor       =   -2147483639
      BackStyle       =   0
      Caption         =   "Click"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel1 
      Height          =   975
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   1720
      BackColor       =   -2147483639
      BackStyle       =   0
      Caption         =   "1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   255
   End
   Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel7 
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   1920
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      Caption         =   "Visual Basic 6.0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel8 
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   2520
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   873
      Alignment       =   1
      BackColor       =   -2147483639
      BackStyle       =   0
      Caption         =   "Chu7o7ng tri2nh co1 su73 du5ng thu7 vie65n OCX FVUnicodeControl.ocx cu3a: (Nhie62u ta1c gia3)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   12583104
   End
   Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel9 
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2160
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      Alignment       =   2
      Caption         =   "Ha5n su73 du5ng:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483635
   End
   Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel10 
      Height          =   255
      Left            =   1920
      TabIndex        =   11
      Top             =   2160
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   450
      Caption         =   "(D9a6y la2 pha62m me62m mie64n phi1 100%)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel11 
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   1200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      Alignment       =   2
      Caption         =   "Phie6n Ba3n:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483635
   End
   Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel12 
      Height          =   255
      Left            =   1920
      TabIndex        =   13
      Top             =   1200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      Caption         =   "1.2.0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel13 
      Height          =   255
      Left            =   600
      TabIndex        =   15
      Top             =   1680
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      Alignment       =   2
      Caption         =   "Website:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483635
   End
   Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel14 
      Height          =   255
      Left            =   1920
      TabIndex        =   16
      Top             =   1680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      Caption         =   "Http://qts.it.tt"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4080
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   5
      Top             =   1440
      Width           =   255
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" _
      Alias "ShellExecuteA" ( _
      ByVal hwnd As Long, _
      ByVal lpOperation As String, _
      ByVal lpFile As String, _
      ByVal lpParameters As String, _
      ByVal lpDirectory As String, _
      ByVal nShowCmd As Long) As Long

Private Sub cmdExit_Click()
Unload Me
End Sub


Private Sub cmdGotoHomePages_Click()
Shell "explorer http://qts.it.tt"
End Sub

Private Sub Label1_Click()
UniMsgBox ChrW$(&H54) & ChrW$(&HE1) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H67) & ChrW$(&H69) & ChrW$(&H1EA3) & ChrW$(&H3A) & ChrW$(&H20) & ChrW$(&H110) & ChrW$(&H69) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H51) & ChrW$(&H75) & ChrW$(&H61) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H54) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H2D) & ChrW$(&H20) & ChrW$(&H31) & ChrW$(&H32) & ChrW$(&H2F) & ChrW$(&H31) & ChrW$(&H32) & ChrW$(&H2F) & ChrW$(&H31) & ChrW$(&H39) & ChrW$(&H39) & ChrW$(&H33) _
& vbCrLf & vbCrLf _
& ChrW$(&H4C) & ChrW$(&H1EDB) & ChrW$(&H70) & ChrW$(&H20) & ChrW$(&H31) & ChrW$(&H30) & ChrW$(&H54) & ChrW$(&H32) & ChrW$(&H20) & ChrW$(&H54) & ChrW$(&H72) & ChrW$(&H1B0) & ChrW$(&H1EDD) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H54) & ChrW$(&H48) & ChrW$(&H50) & ChrW$(&H54) & ChrW$(&H20) & ChrW$(&H110) & ChrW$(&H1ED3) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H50) & ChrW$(&H68) & ChrW$(&HFA) & ChrW$(&H2E) _
& vbCrLf & vbCrLf & ChrW$(&H4D) & ChrW$(&H1ECD) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H68) & ChrW$(&H1EAF) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H6D) & ChrW$(&H1EAF) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H78) & ChrW$(&H69) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H6C) & ChrW$(&H69) & ChrW$(&HEA) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H68) & ChrW$(&H1EC7) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H72) & ChrW$(&H1EF1) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H69) & ChrW$(&H1EBF) & ChrW$(&H70) & ChrW$(&H3A) & ChrW$(&H20) & ChrW$(&H44) & ChrW$(&H69) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H51) & ChrW$(&H75) & ChrW$(&H61) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H54) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H39) & ChrW$(&H30) & ChrW$(&H40) & ChrW$(&H59) & ChrW$(&H61) & ChrW$(&H68) & ChrW$(&H6F) & ChrW$(&H6F) & ChrW$(&H2E) & ChrW$(&H43) & ChrW$(&H6F) & ChrW$(&H6D) _
 , vbOKOnly, "About Me", Me.hwnd

End Sub
