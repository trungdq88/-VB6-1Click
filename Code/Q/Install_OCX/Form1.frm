VERSION 5.00
Object = "{E8FDD05C-3067-4198-8AEC-1A013A46ABDD}#1.0#0"; "FVUnicodeControl.ocx"
Begin VB.Form frmInstall 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Install OCX"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5730
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   5730
   StartUpPosition =   2  'CenterScreen
   Begin FVUnicodeControl.FVistaUniButton cmd 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   960
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      BackColor       =   -2147483633
      ButtonStyle     =   3
      Caption         =   "Ba81t D9a62u Ca2i D9a85t"
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
   Begin FVUnicodeControl.FVistaUniLabel lbl 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   661
      Alignment       =   1
      BackStyle       =   0
      Caption         =   "Ca2i d9a85t ca1c OCX ca62n thie61t d9e63 cha5y chu7o7ng tri2nh 1Click:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
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
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   661
      BackStyle       =   0
      Caption         =   "Nha61n nu1t Ba81t D9a62u Ca2i D9a85t d9e63 ca2i d9a85t OCX"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
End
Attribute VB_Name = "frmInstall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Click()
lbl.Caption = "Xong, ba6y gio72 ba5n co1 the63 su73 du5ng chu7o7ng tri2nh 1Click"
lbl.ForeColor = vbRed
cmd.Enabled = False
End Sub

