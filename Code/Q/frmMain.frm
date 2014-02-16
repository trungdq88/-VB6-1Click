VERSION 5.00
Object = "{E8FDD05C-3067-4198-8AEC-1A013A46ABDD}#1.0#0"; "FVUnicodeControl.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "1Click - 1.2.0"
   ClientHeight    =   8025
   ClientLeft      =   465
   ClientTop       =   540
   ClientWidth     =   10560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   10560
   StartUpPosition =   2  'CenterScreen
   Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel2 
      Height          =   255
      Left            =   4200
      TabIndex        =   164
      Top             =   960
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   450
      Caption         =   "Nhanh Cho1ng, D9o7n Gia3n, De64 Da2ng..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   33023
   End
   Begin FVUnicodeControl.FVistaUniFrame fmCaiDatChung 
      Height          =   4815
      Left            =   3720
      TabIndex        =   103
      Top             =   7440
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   8493
      Alignment       =   0
      BackColor       =   -2147483643
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Ca2i D9a85t Chung"
      AutoUnicode     =   -1  'True
      Begin FVUnicodeControl.FVistaUniLabel lblInternet 
         Height          =   255
         Left            =   2640
         TabIndex        =   136
         Top             =   3840
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   450
         Caption         =   ""
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
      Begin FVUnicodeControl.FVistaUniLabel lblTimeUse 
         Height          =   255
         Left            =   2640
         TabIndex        =   135
         Top             =   3600
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   450
         Caption         =   ""
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
      Begin FVUnicodeControl.FVistaUniLabel lblVirusView 
         Height          =   255
         Left            =   2640
         TabIndex        =   134
         Top             =   3360
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   450
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
      End
      Begin FVUnicodeControl.FVistaUniLabel lblComputerHeal 
         Height          =   255
         Left            =   2640
         TabIndex        =   133
         Top             =   3120
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   450
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
      End
      Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel29 
         Height          =   255
         Left            =   120
         TabIndex        =   132
         Top             =   240
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   450
         Alignment       =   1
         Caption         =   "Nha61n Va2o Nu1t ""D9o1ng Ca2i D9a85t Chung"" D9e63 Tro73 Ve62 Giao Die65n Chi1nh Cu3a Chu7o7ng Tri2nh."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   33023
      End
      Begin FVUnicodeControl.FVistaUniLabel lblLogOnErr 
         Height          =   255
         Left            =   2640
         TabIndex        =   131
         Top             =   2880
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   450
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
      End
      Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel28 
         Height          =   255
         Left            =   2640
         TabIndex        =   130
         Top             =   2640
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   450
         Caption         =   "?"
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
      Begin FVUnicodeControl.FVistaUniLabel lblCheckReg 
         Height          =   255
         Left            =   2640
         TabIndex        =   129
         Top             =   2400
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   450
         Caption         =   ""
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
      Begin FVUnicodeControl.FVistaUniLabel lblCheckTask 
         Height          =   255
         Left            =   2640
         TabIndex        =   128
         Top             =   2160
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   450
         Caption         =   ""
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
      Begin FVUnicodeControl.FVistaUniLabel lblViewAutorun 
         Height          =   255
         Left            =   2640
         TabIndex        =   127
         Top             =   1920
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   450
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
      End
      Begin FVUnicodeControl.FVistaUniLabel lblAutorunPro 
         Height          =   255
         Left            =   2640
         TabIndex        =   126
         Top             =   1680
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   450
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
      End
      Begin FVUnicodeControl.FVistaUniLabel lblSizeRAM 
         Height          =   255
         Left            =   2640
         TabIndex        =   125
         Top             =   1440
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   450
         Caption         =   ""
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
      Begin FVUnicodeControl.FVistaUniLabel lblSizeHardDisk 
         Height          =   255
         Left            =   2640
         TabIndex        =   124
         Top             =   1200
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   450
         Caption         =   ""
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
      Begin FVUnicodeControl.FVistaUniLabel lblComputer 
         Height          =   255
         Left            =   2640
         TabIndex        =   123
         Top             =   960
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   450
         AutoUnicode     =   0   'False
         Caption         =   ""
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
      Begin FVUnicodeControl.FVistaUniLabel lblUser 
         Height          =   255
         Left            =   2640
         TabIndex        =   122
         Top             =   720
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   450
         AutoUnicode     =   0   'False
         Caption         =   ""
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
      Begin FVUnicodeControl.FVistaUniLabel lblHDH 
         Height          =   255
         Left            =   2640
         TabIndex        =   121
         Top             =   480
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   450
         AutoUnicode     =   0   'False
         Caption         =   ""
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
      Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel12 
         Height          =   255
         Left            =   240
         TabIndex        =   106
         Top             =   480
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         Alignment       =   2
         BackStyle       =   0
         Caption         =   "He65 D9ie62u Ha2nh D9ang Su73 Du5ng:"
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
      Begin FVUnicodeControl.FVistaUniButton FVistaUniButton14 
         Height          =   375
         Left            =   8160
         TabIndex        =   104
         Top             =   4320
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BackColor       =   8421631
         ButtonShape     =   2
         Caption         =   "D9o1ng Ca2i D9a85t Chung"
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
      Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel13 
         Height          =   255
         Left            =   240
         TabIndex        =   107
         Top             =   720
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         Alignment       =   2
         BackStyle       =   0
         Caption         =   "Te6n Ngu7o72i Du2ng:"
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
      Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel14 
         Height          =   255
         Left            =   240
         TabIndex        =   108
         Top             =   960
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         Alignment       =   2
         BackStyle       =   0
         Caption         =   "Te6n Ma1y Ti1nh:"
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
      Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel15 
         Height          =   255
         Left            =   240
         TabIndex        =   109
         Top             =   1200
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         Alignment       =   2
         BackStyle       =   0
         Caption         =   "Dung Lu7o75ng O63 Cu71ng:"
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
      Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel16 
         Height          =   255
         Left            =   240
         TabIndex        =   110
         Top             =   1440
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         Alignment       =   2
         BackStyle       =   0
         Caption         =   "Dung Lu7o75ng RAM:"
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
      Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel17 
         Height          =   255
         Left            =   240
         TabIndex        =   111
         Top             =   1680
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         Alignment       =   2
         BackStyle       =   0
         Caption         =   "Ba3o Ve65 Autorun:"
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
      Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel18 
         Height          =   255
         Left            =   240
         TabIndex        =   112
         Top             =   1920
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         Alignment       =   2
         BackStyle       =   0
         Caption         =   "Pha1t Hie65n Autorun:"
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
      Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel19 
         Height          =   255
         Left            =   240
         TabIndex        =   113
         Top             =   2160
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         Alignment       =   2
         BackStyle       =   0
         Caption         =   "Task Manager:"
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
      Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel20 
         Height          =   255
         Left            =   240
         TabIndex        =   114
         Top             =   2400
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         Alignment       =   2
         BackStyle       =   0
         Caption         =   "Registry:"
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
      Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel21 
         Height          =   255
         Left            =   120
         TabIndex        =   115
         Top             =   2640
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
         Alignment       =   2
         BackStyle       =   0
         Caption         =   "Ca1c Chu71c Na8ng Windows Kha1c:"
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
      Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel22 
         Height          =   255
         Left            =   240
         TabIndex        =   116
         Top             =   2880
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         Alignment       =   2
         BackStyle       =   0
         Caption         =   "Lo64i Kho6ng The63 Log On:"
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
      Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel23 
         Height          =   255
         Left            =   240
         TabIndex        =   117
         Top             =   3120
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         Alignment       =   2
         BackStyle       =   0
         Caption         =   "Ti2nh Tra5ng Ma1y Ti1nh:"
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
      Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel24 
         Height          =   255
         Left            =   240
         TabIndex        =   118
         Top             =   3360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         Alignment       =   2
         BackStyle       =   0
         Caption         =   "Pha1t Hie65n Co1 Virus:"
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
      Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel25 
         Height          =   255
         Left            =   240
         TabIndex        =   119
         Top             =   3600
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         Alignment       =   2
         BackStyle       =   0
         Caption         =   "Tho72i Gian Su73 Du5ng Ma1y Ti1nh:"
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
      Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel26 
         Height          =   255
         Left            =   240
         TabIndex        =   120
         Top             =   3840
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         Alignment       =   2
         BackStyle       =   0
         Caption         =   "Ke61t no61i Internet:"
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
   End
   Begin FVUnicodeControl.FVistaUniLabel lblMe 
      Height          =   255
      Left            =   5400
      TabIndex        =   97
      Top             =   0
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   450
      Alignment       =   2
      BackStyle       =   0
      Caption         =   "Copyright © QuangTrung All Right Reserved"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483636
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2640
      Top             =   120
   End
   Begin FVUnicodeControl.FVistaUniLabel lblUpdate 
      Height          =   255
      Left            =   7920
      TabIndex        =   92
      Top             =   1080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      Caption         =   "Ca65p Nha65t"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin FVUnicodeControl.FVistaUniButton cmdMoRong 
      Height          =   255
      Left            =   3840
      TabIndex        =   52
      Top             =   6600
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
      BackColor       =   16761024
      ButtonStyle     =   3
      Caption         =   "Thu Nho3"
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
   Begin FVUnicodeControl.FVistaUniFrame fraHelp 
      Height          =   1215
      Left            =   120
      TabIndex        =   50
      Top             =   6720
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   2143
      Alignment       =   0
      BackColor       =   -2147483643
      ForeColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Tho6ng Tin && Hu7o71ng Da64n"
      AutoUnicode     =   -1  'True
      Begin FVUnicodeControl.FVistaUniLabel lblHelp 
         Height          =   855
         Left            =   120
         TabIndex        =   51
         Top             =   240
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   1508
         BackStyle       =   0
         Caption         =   ""
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
   End
   Begin FVUnicodeControl.FVistaUniTabStrip FVistaUniTabStrip1 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   8916
      TabCount        =   6
      TabCaption(0)   =   "Chu71c Na8ng"
      TabCaption(1)   =   "Phu5c Ho62i"
      TabCaption(2)   =   "Die65t Virus"
      TabCaption(3)   =   "Kie63m Tra He65 Tho61ng"
      TabCaption(4)   =   "Qua3n Ly1 File"
      TabCaption(5)   =   "Ti2m Kie61m"
      AutoUni         =   -1  'True
      ActiveTabBackEndColor=   16777215
      ActiveTabBackStartColor=   16777215
      BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ActiveTabForeColor=   0
      BackColor       =   16777215
      BottomRightInnerBorderColor=   -2147483631
      DisabledTabBackColor=   13355721
      DisabledTabForeColor=   10526880
      InActiveTabBackEndColor=   13619151
      InActiveTabBackStartColor=   15461355
      BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      InActiveTabForeColor=   0
      OuterBorderColor=   9800841
      PictureAlign    =   2
      PictureSize     =   1
      TabStripBackColor=   -2147483639
      UseMouseWheelScroll=   0   'False
      TabOffset       =   11665
      Begin FVUnicodeControl.FVistaUniButton cmdChangePicture 
         Height          =   255
         Left            =   -5545
         TabIndex        =   173
         Top             =   4080
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         BackColor       =   16751432
         Caption         =   "D9o63i Hi2nh"
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
      Begin FVUnicodeControl.FVistaUniButton cmdChangeInfo 
         Height          =   375
         Left            =   -3145
         TabIndex        =   172
         Top             =   4080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BackColor       =   16751432
         Caption         =   "Thay D9o63i Tho6ng Tin"
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
      Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel11 
         Height          =   255
         Left            =   -8425
         TabIndex        =   171
         Top             =   2400
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   450
         BackStyle       =   0
         Caption         =   "Logo He65 Tho61ng:"
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
      Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel10 
         Height          =   255
         Left            =   -5545
         TabIndex        =   170
         Top             =   960
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   450
         BackStyle       =   0
         Caption         =   "Tho6ng Tin He65 Tho61ng:"
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
      Begin FVUnicodeControl.FVistaUniTextbox txtOemInfo 
         Height          =   2775
         Left            =   -5545
         TabIndex        =   169
         Top             =   1200
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   4895
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         Text            =   ""
         MultiLine       =   -1  'True
         BorderStyle     =   2
         BorderLine      =   11709605
         Scrollbar       =   2
      End
      Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel4 
         Height          =   375
         Left            =   -5425
         TabIndex        =   168
         Top             =   600
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
         Alignment       =   1
         Caption         =   "D9o63i Logo + Tho6ng Tin He65 Tho61ng"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
      End
      Begin VB.PictureBox PicOemLogo 
         BackColor       =   &H80000005&
         Height          =   1695
         Left            =   -8425
         ScaleHeight     =   109
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   181
         TabIndex        =   167
         Top             =   2640
         Width           =   2775
      End
      Begin FVUnicodeControl.FVistaUniButton cmdTimKiemFile 
         Height          =   375
         Left            =   -39700
         TabIndex        =   150
         Top             =   1080
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BackColor       =   -2147483646
         ButtonShape     =   3
         ButtonStyle     =   1
         Caption         =   "Ti2m Kie61m File"
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
      Begin FVUnicodeControl.FVistaUniButton cmdSaveLog 
         Height          =   375
         Left            =   -25995
         TabIndex        =   102
         Top             =   4440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BackColor       =   -2147483646
         ButtonShape     =   3
         ButtonStyle     =   1
         Caption         =   "Lu7u la5i Log file"
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
      Begin FVUnicodeControl.FVistaUniButton cmdKiemtra 
         Height          =   375
         Left            =   -31155
         TabIndex        =   101
         Top             =   4440
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         BackColor       =   -2147483646
         ButtonShape     =   3
         ButtonStyle     =   1
         Caption         =   "Kie63m Tra He65 Tho61ng"
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
      Begin FVUnicodeControl.FVistaUniTextbox txtCheck 
         Height          =   3735
         Left            =   -34875
         TabIndex        =   100
         Top             =   480
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   6588
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         Text            =   ""
         MultiLine       =   -1  'True
         Locked          =   -1  'True
         BorderStyle     =   2
         BorderLine      =   11709605
         Scrollbar       =   3
      End
      Begin FVUnicodeControl.FVistaUniButton FVistaUniButton10 
         Height          =   495
         Left            =   -11545
         TabIndex        =   95
         Top             =   4080
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   873
         BackColor       =   -2147483635
         ButtonShape     =   3
         Caption         =   "Su73a lo64i kho6ng the63 cha5y File EXE"
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
      Begin FVUnicodeControl.FVistaUniButton FVistaUniButton6 
         Height          =   495
         Left            =   -11545
         TabIndex        =   81
         Top             =   3480
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   873
         BackColor       =   -2147483635
         ButtonShape     =   3
         Caption         =   "Su73a Lo64i Kho6ng The63 Log On"
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
      Begin FVUnicodeControl.FVistaUniFrame FVistaUniFrame6 
         Height          =   4455
         Left            =   -17450
         TabIndex        =   80
         Top             =   480
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   7858
         Alignment       =   0
         BackColor       =   -2147483643
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Xo1a File Kho6ng The63 Xo1a D9u7o75c"
         AutoUnicode     =   -1  'True
         Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel9 
            Height          =   855
            Left            =   1080
            TabIndex        =   90
            Top             =   3360
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   1508
            Caption         =   "Nha61n nu1t The6m d9e63 the6m File va2o danh sa1ch xo1a. Sau d9o1 nha61n nu1t Xa1c Nha65n d9e63 xo1a File"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   255
         End
         Begin VB.ListBox List1 
            Height          =   2010
            ItemData        =   "frmMain.frx":57E2
            Left            =   120
            List            =   "frmMain.frx":57E4
            TabIndex        =   89
            Top             =   240
            Width           =   4095
         End
         Begin FVUnicodeControl.FVistaUniLabel lblFile 
            Height          =   375
            Left            =   240
            TabIndex        =   88
            Top             =   2280
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   661
            BackStyle       =   0
            Caption         =   ""
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
         Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel8 
            Height          =   615
            Left            =   120
            TabIndex        =   87
            Top             =   2640
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   1085
            BackStyle       =   0
            Caption         =   $"frmMain.frx":57E6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   255
         End
         Begin FVUnicodeControl.FVistaUniButton cmdCancel 
            Height          =   375
            Left            =   3360
            TabIndex        =   86
            Top             =   3840
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            BackColor       =   8421631
            ButtonShape     =   3
            ButtonStyle     =   1
            Caption         =   "Hu3y Bo3"
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
         Begin FVUnicodeControl.FVistaUniButton cmdOK 
            Height          =   375
            Left            =   3360
            TabIndex        =   85
            Top             =   3360
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            BackColor       =   8421631
            ButtonShape     =   3
            ButtonStyle     =   1
            Caption         =   "Xa1c Nha65n"
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
         Begin FVUnicodeControl.FVistaUniButton cmdDelFile 
            Height          =   375
            Left            =   120
            TabIndex        =   84
            Top             =   3840
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            BackColor       =   8421631
            ButtonShape     =   3
            ButtonStyle     =   1
            Caption         =   "Xo1a"
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
         Begin FVUnicodeControl.FVistaUniButton cmdAddFile 
            Height          =   375
            Left            =   120
            TabIndex        =   83
            Top             =   3360
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            BackColor       =   8421631
            ButtonShape     =   3
            ButtonStyle     =   1
            Caption         =   "The6m"
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
      End
      Begin FVUnicodeControl.FVistaUniButton FVistaUniButton5 
         Height          =   495
         Left            =   -11545
         TabIndex        =   49
         Top             =   2880
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   873
         BackColor       =   -2147483635
         ButtonShape     =   3
         Caption         =   "Su73a Chu74a Lo64i Chuye63n hu7o71ng Trang Web"
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
      Begin FVUnicodeControl.FVistaUniFrame FVistaUniFrame5 
         Height          =   4455
         Left            =   -20330
         TabIndex        =   41
         Top             =   480
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   7858
         Alignment       =   0
         BackColor       =   -2147483643
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Nga8n Kho6ng Cho Mo65t File Cha5y"
         AutoUnicode     =   -1  'True
         Begin FVUnicodeControl.FVistaUniButton FVistaUniButton9 
            Height          =   255
            Left            =   2280
            TabIndex        =   82
            Top             =   600
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   450
            BackColor       =   -2147483635
            ButtonShape     =   3
            ButtonStyle     =   1
            Caption         =   "..."
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
         Begin FVUnicodeControl.FVistaUniButton FVistaUniButton8 
            Height          =   225
            Left            =   1200
            TabIndex        =   48
            Top             =   4080
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   397
            BackColor       =   -2147483635
            ButtonShape     =   3
            ButtonStyle     =   1
            Caption         =   "Thu75c Hie65n"
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
         Begin FVUnicodeControl.FVistaUniButton cmdDel 
            Height          =   225
            Left            =   120
            TabIndex        =   47
            Top             =   4080
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   397
            BackColor       =   -2147483635
            ButtonShape     =   3
            ButtonStyle     =   1
            Caption         =   "Xoa1"
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
         Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel7 
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   1320
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   450
            BackStyle       =   0
            Caption         =   "Danh Sa1ch Ca1c File Bi5 Ca61m"
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
         Begin FVUnicodeControl.FVistaUniListbox ListCam 
            Height          =   2400
            Left            =   120
            TabIndex        =   45
            Top             =   1560
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   4233
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ItemHeight      =   13
            AutoUni         =   -1  'True
         End
         Begin FVUnicodeControl.FVistaUniButton cmdAdd 
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
            BackColor       =   -2147483635
            ButtonShape     =   3
            ButtonStyle     =   1
            Caption         =   "Ca61m Cha5y File Na2y"
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
         Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel6 
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   360
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   450
            BackStyle       =   0
            Caption         =   "Cho5n File..."
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
         Begin FVUnicodeControl.FVistaUniTextbox txtFileName 
            Height          =   270
            Left            =   120
            TabIndex        =   42
            Top             =   600
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   476
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   0
            Text            =   ""
            BorderStyle     =   2
            BorderLine      =   11709605
         End
      End
      Begin FVUnicodeControl.FVistaUniFrame FVistaUniFrame4 
         Height          =   4455
         Left            =   -23210
         TabIndex        =   38
         Top             =   480
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   7858
         Alignment       =   0
         BackColor       =   -2147483643
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Die65t Virus - 1Click la2 xong"
         AutoUnicode     =   -1  'True
         Begin OneClick.McListBox ListVR 
            Height          =   3375
            Left            =   120
            TabIndex        =   137
            Top             =   480
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   5953
            Picture         =   "frmMain.frx":58A9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   131072
         End
         Begin FVUnicodeControl.FVistaUniLabel slVirus 
            Height          =   255
            Left            =   2280
            TabIndex        =   94
            Top             =   240
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   450
            BackStyle       =   0
            Caption         =   "(0)"
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
         Begin FVUnicodeControl.FVistaUniButton cmdCheckVR 
            Height          =   375
            Left            =   1680
            TabIndex        =   91
            Top             =   3960
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            BackColor       =   192
            ButtonShape     =   3
            ButtonStyle     =   1
            Caption         =   "Que1t Virus"
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
         Begin FVUnicodeControl.FVistaUniButton cmdDiet 
            Height          =   375
            Left            =   120
            TabIndex        =   39
            Top             =   3960
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            BackColor       =   -2147483635
            ButtonShape     =   3
            ButtonStyle     =   1
            Caption         =   "Die65t Nhanh"
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            State           =   3
         End
         Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel5 
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   450
            BackStyle       =   0
            Caption         =   "Danh Sa1ch Virus Co1 The63 Die65t"
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
      End
      Begin FVUnicodeControl.FVistaUniFrame FVistaUniFrame3 
         Height          =   2175
         Left            =   6840
         TabIndex        =   1
         Top             =   600
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   3836
         Alignment       =   0
         BackColor       =   -2147483643
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Ca1c Ti1nh Na8ng Kha1c"
         AutoUnicode     =   -1  'True
         Begin FVUnicodeControl.FVistaUniCheckbox chkWrite 
            Height          =   195
            Left            =   240
            TabIndex        =   2
            Top             =   1800
            Width           =   1530
            _ExtentX        =   2699
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
            Caption         =   "Ca61m Ghi Va2o USB"
            ForeColor       =   4210688
         End
         Begin FVUnicodeControl.FVistaUniCheckbox chkUSB 
            Height          =   195
            Left            =   240
            TabIndex        =   3
            Top             =   1440
            Width           =   1110
            _ExtentX        =   1958
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
            Caption         =   "Ca61m O63 USB"
            ForeColor       =   4210688
         End
         Begin FVUnicodeControl.FVistaUniCheckbox chkAUTORUN 
            Height          =   195
            Left            =   240
            TabIndex        =   4
            Top             =   1080
            Width           =   2310
            _ExtentX        =   4075
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
            Caption         =   "Kho6ng cha5y Autorun tu72 USB"
            ForeColor       =   4210688
         End
         Begin FVUnicodeControl.FVistaUniCheckbox chkHIDDEN 
            Height          =   195
            Left            =   240
            TabIndex        =   5
            Top             =   720
            Width           =   1890
            _ExtentX        =   3334
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
            Caption         =   "Kho6ng Hie63n Thi5 File A63n"
            ForeColor       =   4210688
         End
         Begin FVUnicodeControl.FVistaUniCheckbox chkEXE 
            Height          =   195
            Left            =   240
            TabIndex        =   6
            Top             =   360
            Width           =   2025
            _ExtentX        =   3572
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
            Caption         =   "Kho6ng Hie63n Thi5 D9uo6i File"
            ForeColor       =   4210688
         End
         Begin FVUnicodeControl.FVistaUniButton EXEHelp 
            Height          =   255
            Left            =   3000
            TabIndex        =   75
            ToolTipText     =   "Help"
            Top             =   360
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            BackColor       =   12632256
            ButtonStyle     =   9
            Caption         =   "?"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   14
            PicOpacity      =   0
         End
         Begin FVUnicodeControl.FVistaUniButton HIDEHelp 
            Height          =   255
            Left            =   3000
            TabIndex        =   76
            ToolTipText     =   "Help"
            Top             =   720
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            BackColor       =   12632256
            ButtonStyle     =   9
            Caption         =   "?"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   14
            PicOpacity      =   0
         End
         Begin FVUnicodeControl.FVistaUniButton AutorunHelp 
            Height          =   255
            Left            =   3000
            TabIndex        =   77
            ToolTipText     =   "Help"
            Top             =   1080
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            BackColor       =   12632256
            ButtonStyle     =   9
            Caption         =   "?"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   14
            PicOpacity      =   0
         End
         Begin FVUnicodeControl.FVistaUniButton USBHelp 
            Height          =   255
            Left            =   3000
            TabIndex        =   78
            ToolTipText     =   "Help"
            Top             =   1440
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            BackColor       =   12632256
            ButtonStyle     =   9
            Caption         =   "?"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   14
            PicOpacity      =   0
         End
         Begin FVUnicodeControl.FVistaUniButton WriteHelp 
            Height          =   255
            Left            =   3000
            TabIndex        =   79
            ToolTipText     =   "Help"
            Top             =   1800
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            BackColor       =   12632256
            ButtonStyle     =   9
            Caption         =   "?"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   14
            PicOpacity      =   0
         End
      End
      Begin FVUnicodeControl.FVistaUniFrame FVistaUniFrame2 
         Height          =   4335
         Left            =   3480
         TabIndex        =   7
         Top             =   600
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   7646
         Alignment       =   0
         BackColor       =   -2147483643
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "A63n/Hie65n Ca1c Ti1nh Na8ng"
         AutoUnicode     =   -1  'True
         Begin FVUnicodeControl.FVistaUniCheckbox chkDOC 
            Height          =   195
            Left            =   240
            TabIndex        =   8
            Top             =   3960
            Width           =   1515
            _ExtentX        =   2672
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
            Caption         =   "A63n Folder Option"
            ForeColor       =   4210688
         End
         Begin FVUnicodeControl.FVistaUniCheckbox chkPRO 
            Height          =   195
            Left            =   240
            TabIndex        =   9
            Top             =   3600
            Width           =   2385
            _ExtentX        =   4207
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
            Caption         =   "A63n All Programs (Start Menu)"
            ForeColor       =   4210688
         End
         Begin FVUnicodeControl.FVistaUniCheckbox chkTURNOFF 
            Height          =   195
            Left            =   240
            TabIndex        =   10
            Top             =   360
            Width           =   2415
            _ExtentX        =   4260
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
            Caption         =   "A63n Nu1t Turn Off (Start Menu)"
            ForeColor       =   4210688
         End
         Begin FVUnicodeControl.FVistaUniCheckbox chkLOGOFF 
            Height          =   195
            Left            =   240
            TabIndex        =   11
            Top             =   720
            Width           =   2340
            _ExtentX        =   4128
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
            Caption         =   "A63n Nu1t Log Off (Start Menu)"
            ForeColor       =   4210688
         End
         Begin FVUnicodeControl.FVistaUniCheckbox chkSearch 
            Height          =   195
            Left            =   240
            TabIndex        =   12
            Top             =   1080
            Width           =   2520
            _ExtentX        =   4445
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
            Caption         =   "A63n Search Engine (Start Menu)"
            ForeColor       =   4210688
         End
         Begin FVUnicodeControl.FVistaUniCheckbox chkHelp 
            Height          =   195
            Left            =   240
            TabIndex        =   13
            Top             =   1440
            Width           =   1785
            _ExtentX        =   3149
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
            Caption         =   "A63n Help and Support"
            ForeColor       =   4210688
         End
         Begin FVUnicodeControl.FVistaUniCheckbox chkRUN 
            Height          =   195
            Left            =   240
            TabIndex        =   14
            Top             =   1800
            Width           =   2145
            _ExtentX        =   3784
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
            Caption         =   "A63n Tri2nh Le65nh D9o7n Run..."
            ForeColor       =   4210688
         End
         Begin FVUnicodeControl.FVistaUniCheckbox chkCPITEM 
            Height          =   195
            Left            =   240
            TabIndex        =   15
            Top             =   2160
            Width           =   1950
            _ExtentX        =   3440
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
            Caption         =   "A63n Control Panel Items"
            ForeColor       =   4210688
         End
         Begin FVUnicodeControl.FVistaUniCheckbox chkTRAYCLOCK 
            Height          =   195
            Left            =   240
            TabIndex        =   16
            Top             =   2520
            Width           =   1980
            _ExtentX        =   3493
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
            Caption         =   "A63n Tray Icons && Clock"
            ForeColor       =   4210688
         End
         Begin FVUnicodeControl.FVistaUniCheckbox chkCOMPUTER 
            Height          =   195
            Left            =   240
            TabIndex        =   17
            Top             =   2880
            Width           =   2025
            _ExtentX        =   3572
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
            Caption         =   "A63n Computer Properties"
            ForeColor       =   4210688
         End
         Begin FVUnicodeControl.FVistaUniCheckbox chkCPA 
            Height          =   195
            Left            =   240
            TabIndex        =   18
            Top             =   3240
            Width           =   1500
            _ExtentX        =   2646
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
            Caption         =   "A63n Control Panel"
            ForeColor       =   4210688
         End
         Begin FVUnicodeControl.FVistaUniButton OffHelp 
            Height          =   255
            Left            =   2880
            TabIndex        =   64
            ToolTipText     =   "Help"
            Top             =   360
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            BackColor       =   12632256
            ButtonStyle     =   9
            Caption         =   "?"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   14
            PicOpacity      =   0
         End
         Begin FVUnicodeControl.FVistaUniButton LogOffHelp 
            Height          =   255
            Left            =   2880
            TabIndex        =   65
            ToolTipText     =   "Help"
            Top             =   720
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            BackColor       =   12632256
            ButtonStyle     =   9
            Caption         =   "?"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   14
            PicOpacity      =   0
         End
         Begin FVUnicodeControl.FVistaUniButton StartHelp 
            Height          =   255
            Left            =   2880
            TabIndex        =   66
            ToolTipText     =   "Help"
            Top             =   1080
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            BackColor       =   12632256
            ButtonStyle     =   9
            Caption         =   "?"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   14
            PicOpacity      =   0
         End
         Begin FVUnicodeControl.FVistaUniButton HelpHelp 
            Height          =   255
            Left            =   2880
            TabIndex        =   67
            ToolTipText     =   "Help"
            Top             =   1440
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            BackColor       =   12632256
            ButtonStyle     =   9
            Caption         =   "?"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   14
            PicOpacity      =   0
         End
         Begin FVUnicodeControl.FVistaUniButton RunHelp 
            Height          =   255
            Left            =   2880
            TabIndex        =   68
            ToolTipText     =   "Help"
            Top             =   1800
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            BackColor       =   12632256
            ButtonStyle     =   9
            Caption         =   "?"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   14
            PicOpacity      =   0
         End
         Begin FVUnicodeControl.FVistaUniButton CPITEMHelp 
            Height          =   255
            Left            =   2880
            TabIndex        =   69
            ToolTipText     =   "Help"
            Top             =   2160
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            BackColor       =   12632256
            ButtonStyle     =   9
            Caption         =   "?"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   14
            PicOpacity      =   0
         End
         Begin FVUnicodeControl.FVistaUniButton ClockHelp 
            Height          =   255
            Left            =   2880
            TabIndex        =   70
            ToolTipText     =   "Help"
            Top             =   2520
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            BackColor       =   12632256
            ButtonStyle     =   9
            Caption         =   "?"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   14
            PicOpacity      =   0
         End
         Begin FVUnicodeControl.FVistaUniButton ComputerHelp 
            Height          =   255
            Left            =   2880
            TabIndex        =   71
            ToolTipText     =   "Help"
            Top             =   2880
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            BackColor       =   12632256
            ButtonStyle     =   9
            Caption         =   "?"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   14
            PicOpacity      =   0
         End
         Begin FVUnicodeControl.FVistaUniButton HIDECPHelp 
            Height          =   255
            Left            =   2880
            TabIndex        =   72
            ToolTipText     =   "Help"
            Top             =   3240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            BackColor       =   12632256
            ButtonStyle     =   9
            Caption         =   "?"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   14
            PicOpacity      =   0
         End
         Begin FVUnicodeControl.FVistaUniButton ProgramHelp 
            Height          =   255
            Left            =   2880
            TabIndex        =   73
            ToolTipText     =   "Help"
            Top             =   3600
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            BackColor       =   12632256
            ButtonStyle     =   9
            Caption         =   "?"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   14
            PicOpacity      =   0
         End
         Begin FVUnicodeControl.FVistaUniButton DocHelp 
            Height          =   255
            Left            =   2880
            TabIndex        =   74
            ToolTipText     =   "Help"
            Top             =   3960
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            BackColor       =   12632256
            ButtonStyle     =   9
            Caption         =   "?"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   14
            PicOpacity      =   0
         End
      End
      Begin FVUnicodeControl.FVistaUniFrame FVistaUniFrame1 
         Height          =   4335
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   7646
         Alignment       =   0
         BackColor       =   -2147483643
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Kho1a/Mo73 Ca1c Ti1nh Na8ng"
         AutoUnicode     =   -1  'True
         Begin FVUnicodeControl.FVistaUniCheckbox chkWin 
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   3960
            Width           =   2325
            _ExtentX        =   4101
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
            Caption         =   "Kho1a Phi1m Windows + [Key]"
            ForeColor       =   4210688
         End
         Begin FVUnicodeControl.FVistaUniCheckbox chkTaskbar 
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   3600
            Width           =   2730
            _ExtentX        =   4815
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
            Caption         =   "Ko Cho Ca2i D9a85t Taskbar Va2 Folder"
            ForeColor       =   4210688
         End
         Begin FVUnicodeControl.FVistaUniCheckbox chkDESKTOP 
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   3240
            Width           =   2730
            _ExtentX        =   4815
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
            Caption         =   "Kho6ng Hie63n Thi5 Icon Tre6n Desktop"
            ForeColor       =   4210688
         End
         Begin FVUnicodeControl.FVistaUniCheckbox chkFILEMENU 
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   2880
            Width           =   2145
            _ExtentX        =   3784
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
            Caption         =   "Kho1a File Menu (Explorer)"
            ForeColor       =   4210688
         End
         Begin FVUnicodeControl.FVistaUniCheckbox chkRIGHT 
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   2520
            Width           =   1920
            _ExtentX        =   3387
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
            Caption         =   "Kho1a Menu Chuo65t Pha3i"
            ForeColor       =   4210688
         End
         Begin FVUnicodeControl.FVistaUniCheckbox chkTRAY 
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   2160
            Width           =   2100
            _ExtentX        =   3704
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
            Caption         =   "Kho1a Tray Context Menu"
            ForeColor       =   4210688
         End
         Begin FVUnicodeControl.FVistaUniCheckbox chkIEHOME 
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   1800
            Width           =   2475
            _ExtentX        =   4366
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
            Caption         =   "Kho1a Thay D9o63i IE Home Pages"
            ForeColor       =   4210688
         End
         Begin FVUnicodeControl.FVistaUniCheckbox chkCP 
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   1440
            Width           =   1665
            _ExtentX        =   2937
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
            Caption         =   "Kho1a Control Panel"
            ForeColor       =   4210688
         End
         Begin FVUnicodeControl.FVistaUniCheckbox chkCMD 
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   1080
            Width           =   2460
            _ExtentX        =   4339
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
            Caption         =   "Kho1a Command Prompt (CMD)"
            ForeColor       =   4210688
         End
         Begin FVUnicodeControl.FVistaUniCheckbox chkREG 
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   720
            Width           =   1770
            _ExtentX        =   3122
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
            Caption         =   "Kho1a Registry Editor"
            ForeColor       =   4210688
         End
         Begin FVUnicodeControl.FVistaUniCheckbox chkTask 
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   360
            Width           =   2400
            _ExtentX        =   4233
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
            Caption         =   "Kho1a Windows Task Manager"
            ForeColor       =   4210688
         End
         Begin FVUnicodeControl.FVistaUniButton TaskHelp 
            Height          =   255
            Left            =   2880
            TabIndex        =   53
            ToolTipText     =   "Help"
            Top             =   360
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            BackColor       =   12632256
            ButtonShape     =   3
            ButtonStyle     =   9
            Caption         =   "?"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   14
            PicOpacity      =   0
         End
         Begin FVUnicodeControl.FVistaUniButton RegHelp 
            Height          =   255
            Left            =   2880
            TabIndex        =   54
            ToolTipText     =   "Help"
            Top             =   720
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            BackColor       =   12632256
            ButtonShape     =   3
            ButtonStyle     =   9
            Caption         =   "?"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   14
            PicOpacity      =   0
         End
         Begin FVUnicodeControl.FVistaUniButton ComHelp 
            Height          =   255
            Left            =   2880
            TabIndex        =   55
            ToolTipText     =   "Help"
            Top             =   1080
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            BackColor       =   12632256
            ButtonShape     =   3
            ButtonStyle     =   9
            Caption         =   "?"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   14
            PicOpacity      =   0
         End
         Begin FVUnicodeControl.FVistaUniButton CPHelp 
            Height          =   255
            Left            =   2880
            TabIndex        =   56
            ToolTipText     =   "Help"
            Top             =   1440
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            BackColor       =   12632256
            ButtonShape     =   3
            ButtonStyle     =   9
            Caption         =   "?"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   14
            PicOpacity      =   0
         End
         Begin FVUnicodeControl.FVistaUniButton IEHelp 
            Height          =   255
            Left            =   2880
            TabIndex        =   57
            ToolTipText     =   "Help"
            Top             =   1800
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            BackColor       =   12632256
            ButtonShape     =   3
            ButtonStyle     =   9
            Caption         =   "?"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   14
            PicOpacity      =   0
         End
         Begin FVUnicodeControl.FVistaUniButton TrayHelp 
            Height          =   255
            Left            =   2880
            TabIndex        =   58
            ToolTipText     =   "Help"
            Top             =   2160
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            BackColor       =   12632256
            ButtonShape     =   3
            ButtonStyle     =   9
            Caption         =   "?"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   14
            PicOpacity      =   0
         End
         Begin FVUnicodeControl.FVistaUniButton RIGHTHelp 
            Height          =   255
            Left            =   2880
            TabIndex        =   59
            ToolTipText     =   "Help"
            Top             =   2520
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            BackColor       =   12632256
            ButtonShape     =   3
            ButtonStyle     =   9
            Caption         =   "?"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   14
            PicOpacity      =   0
         End
         Begin FVUnicodeControl.FVistaUniButton MENUHelp 
            Height          =   255
            Left            =   2880
            TabIndex        =   60
            ToolTipText     =   "Help"
            Top             =   2880
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            BackColor       =   12632256
            ButtonShape     =   3
            ButtonStyle     =   9
            Caption         =   "?"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   14
            PicOpacity      =   0
         End
         Begin FVUnicodeControl.FVistaUniButton DeskHelp 
            Height          =   255
            Left            =   2880
            TabIndex        =   61
            ToolTipText     =   "Help"
            Top             =   3240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            BackColor       =   12632256
            ButtonShape     =   3
            ButtonStyle     =   9
            Caption         =   "?"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   14
            PicOpacity      =   0
         End
         Begin FVUnicodeControl.FVistaUniButton TaskbarHelp 
            Height          =   255
            Left            =   2880
            TabIndex        =   62
            ToolTipText     =   "Help"
            Top             =   3600
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            BackColor       =   12632256
            ButtonShape     =   3
            ButtonStyle     =   9
            Caption         =   "?"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   14
            PicOpacity      =   0
         End
         Begin FVUnicodeControl.FVistaUniButton WinHelp 
            Height          =   255
            Left            =   2880
            TabIndex        =   63
            ToolTipText     =   "Help"
            Top             =   3960
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            BackColor       =   12632256
            ButtonShape     =   3
            ButtonStyle     =   9
            Caption         =   "?"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   14
            PicOpacity      =   0
         End
      End
      Begin FVUnicodeControl.FVistaUniButton cmdSet 
         Height          =   375
         Left            =   8520
         TabIndex        =   31
         Top             =   4560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BackColor       =   12632256
         ButtonStyle     =   3
         Caption         =   "Thu75c Hie65n"
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
         HandPointer     =   -1  'True
         PicNormal       =   "frmMain.frx":58C5
         PicSize         =   1
         PicSizeH        =   16
         PicSizeW        =   16
      End
      Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel1 
         Height          =   855
         Left            =   7080
         TabIndex        =   32
         Top             =   2880
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   1508
         Caption         =   $"frmMain.frx":1AA37
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   255
      End
      Begin FVUnicodeControl.FVistaUniButton FVistaUniButton7 
         Height          =   495
         Left            =   -9145
         TabIndex        =   33
         Top             =   1680
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         BackColor       =   -2147483635
         ButtonShape     =   3
         Caption         =   "Go74"
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
      Begin FVUnicodeControl.FVistaUniButton FVistaUniButton3 
         Height          =   495
         Left            =   -11545
         TabIndex        =   34
         Top             =   1680
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
         BackColor       =   -2147483635
         ButtonShape     =   3
         Caption         =   "Ba3o Ve65 Ma1y Ti1nh Tra1nh Kho3i Autorun"
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
      Begin FVUnicodeControl.FVistaUniButton FVistaUniButton1 
         Height          =   495
         Left            =   -11545
         TabIndex        =   35
         Top             =   480
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   873
         BackColor       =   -2147483635
         ButtonShape     =   3
         Caption         =   "Xo1a Toa2n Bo65 Auturun.inf"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicAlign        =   3
         PicOpacity      =   0
      End
      Begin FVUnicodeControl.FVistaUniButton FVistaUniButton4 
         Height          =   495
         Left            =   -11545
         TabIndex        =   36
         Top             =   2280
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   873
         BackColor       =   -2147483635
         ButtonShape     =   3
         Caption         =   "Ta8ng To61c Ma1y Ti1nh"
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
      Begin FVUnicodeControl.FVistaUniButton FVistaUniButton2 
         Height          =   495
         Left            =   -11545
         TabIndex        =   37
         Top             =   1080
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   873
         BackColor       =   -2147483635
         ButtonShape     =   3
         Caption         =   "Phu5c Ho62i Ma85c D9i5nh Internet Explorer"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicAlign        =   3
         PicOpacity      =   0
      End
      Begin FVUnicodeControl.FVistaUniButton FVistaUniButton12 
         Height          =   495
         Left            =   -8545
         TabIndex        =   98
         Top             =   480
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   873
         BackColor       =   -2147483635
         ButtonShape     =   3
         Caption         =   "D9a8ng Ky1 Ba3n Quye62n Windows XP"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicAlign        =   3
         PicOpacity      =   0
      End
      Begin FVUnicodeControl.FVistaUniButton FVistaUniButton13 
         Height          =   495
         Left            =   -8545
         TabIndex        =   99
         Top             =   1080
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   873
         BackColor       =   -2147483635
         ButtonShape     =   3
         Caption         =   "Hie65n Ta61t Ca3 Ca1c O63 D9i4a"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicAlign        =   3
         PicOpacity      =   0
      End
      Begin FVUnicodeControl.FVistaUniFrame fmSearch 
         Height          =   4575
         Left            =   -58205
         TabIndex        =   138
         Top             =   360
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   8070
         Alignment       =   0
         BackColor       =   16777215
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Ti2m Kie61m"
         AutoUnicode     =   -1  'True
         Begin FVUnicodeControl.FVistaUniLabel lblInfoSize 
            Height          =   255
            Left            =   8160
            TabIndex        =   188
            Top             =   2880
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
            Caption         =   ""
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
         Begin VB.PictureBox PicInfoIcon 
            BackColor       =   &H80000005&
            Height          =   735
            Left            =   8640
            ScaleHeight     =   675
            ScaleWidth      =   675
            TabIndex        =   187
            Top             =   3480
            Width           =   735
         End
         Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel31 
            Height          =   255
            Left            =   8160
            TabIndex        =   186
            Top             =   3240
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            Caption         =   "Bie63u Tu7o75ng:"
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
         Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel30 
            Height          =   255
            Left            =   8160
            TabIndex        =   185
            Top             =   2640
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            Caption         =   "Ki1ch Thu7o71c:"
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
         Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel27 
            Height          =   255
            Left            =   8160
            TabIndex        =   184
            Top             =   2400
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
            Caption         =   "Tho6ng Tin File:"
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
         Begin FVUnicodeControl.FVistaUniButton cmdStopSearch 
            Height          =   255
            Left            =   7080
            TabIndex        =   182
            Top             =   4200
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            BackColor       =   16751432
            ButtonShape     =   3
            Caption         =   "Du72ng"
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
         Begin VB.PictureBox PicIcon2 
            BackColor       =   &H80000014&
            Height          =   615
            Left            =   9120
            ScaleHeight     =   555
            ScaleWidth      =   555
            TabIndex        =   174
            Top             =   -240
            Visible         =   0   'False
            Width           =   615
         End
         Begin FVUnicodeControl.FVistaUniFrame fmNangCao 
            Height          =   3495
            Left            =   720
            TabIndex        =   153
            Top             =   480
            Visible         =   0   'False
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   6165
            Alignment       =   0
            BackColor       =   16761024
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Ti2m Kie61m Na6ng Cao"
            AutoUnicode     =   -1  'True
            Begin FVUnicodeControl.FVistaUniOption optSearchByIcon 
               Height          =   195
               Left            =   360
               TabIndex        =   175
               Top             =   1320
               Width           =   2655
               _ExtentX        =   4683
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
               Caption         =   "Ti2m Kie61m Theo Bie63u Tu7o75ng (Icon)"
               BackStyle       =   0
               ForeColor       =   0
            End
            Begin FVUnicodeControl.FVistaUniFrame fmSearchTheoIcon 
               Height          =   1215
               Left            =   600
               TabIndex        =   176
               Top             =   1560
               Width           =   4815
               _ExtentX        =   8493
               _ExtentY        =   2143
               Alignment       =   0
               BackColor       =   16761024
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               AutoUnicode     =   0   'False
               Begin VB.PictureBox PicIcon1 
                  BackColor       =   &H80000005&
                  Enabled         =   0   'False
                  Height          =   615
                  Left            =   120
                  MousePointer    =   2  'Cross
                  ScaleHeight     =   555
                  ScaleWidth      =   555
                  TabIndex        =   177
                  Top             =   480
                  Width           =   615
               End
               Begin FVUnicodeControl.FVistaUniOption chkIconChinhXac 
                  Height          =   195
                  Left            =   840
                  TabIndex        =   178
                  Top             =   480
                  Width           =   3420
                  _ExtentX        =   6033
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
                  Caption         =   "Ti2m Kie61m Chi1nh Xa1c (Gio61ng Nhau Tre6n 80%)"
                  Enabled         =   0   'False
                  BackStyle       =   0
                  BackColor       =   -2147483639
                  ForeColor       =   0
               End
               Begin FVUnicodeControl.FVistaUniOption chkIconGanGiong 
                  Height          =   195
                  Left            =   840
                  TabIndex        =   179
                  Top             =   720
                  Width           =   3450
                  _ExtentX        =   6085
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
                  Caption         =   "Ti2m Kie61m Ga62n Gio61ng (Gio61ng Nhau Tre6n 50%)"
                  Enabled         =   0   'False
                  BackStyle       =   0
                  BackColor       =   -2147483639
                  ForeColor       =   0
               End
               Begin FVUnicodeControl.FVistaUniOption chkIconTimHet 
                  Height          =   195
                  Left            =   840
                  TabIndex        =   180
                  Top             =   960
                  Width           =   2820
                  _ExtentX        =   4974
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
                  Caption         =   "Ti2m Ta61t Ca3 (Gio61ng Nhau Tre6n 30%)"
                  Enabled         =   0   'False
                  BackStyle       =   0
                  BackColor       =   -2147483639
                  ForeColor       =   0
               End
               Begin FVUnicodeControl.FVistaUniLabel lblHelpClickIcon 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   181
                  Top             =   240
                  Width           =   4095
                  _ExtentX        =   7223
                  _ExtentY        =   450
                  BackStyle       =   0
                  Caption         =   "(Click chuo65t va2o khung ma2u tra81ng d9e63 cho5n Icon)"
                  Enabled         =   0   'False
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
            End
            Begin FVUnicodeControl.FVistaUniButton FVistaUniButton15 
               Height          =   375
               Left            =   4440
               TabIndex        =   166
               Top             =   840
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   661
               BackColor       =   14737632
               ButtonShape     =   3
               ButtonStyle     =   9
               Caption         =   "?"
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
            Begin FVUnicodeControl.FVistaUniFrame FVistaUniFrame7 
               Height          =   375
               Left            =   120
               TabIndex        =   157
               Top             =   2760
               Width           =   5535
               _ExtentX        =   9763
               _ExtentY        =   661
               Alignment       =   0
               BackColor       =   16761024
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               AutoUnicode     =   0   'False
               Begin FVUnicodeControl.FVistaUniOption optAllDuoi 
                  Height          =   195
                  Left            =   2880
                  TabIndex        =   158
                  Top             =   120
                  Width           =   2490
                  _ExtentX        =   4392
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
                  Caption         =   "*.* (Ta61t Ca3 Ca1c Loa5i D9uo6i File)"
                  Enabled         =   0   'False
                  BackStyle       =   0
                  BackColor       =   -2147483639
                  ForeColor       =   0
               End
               Begin FVUnicodeControl.FVistaUniOption optEXEDuoi 
                  Height          =   195
                  Left            =   1920
                  TabIndex        =   159
                  Top             =   120
                  Width           =   720
                  _ExtentX        =   1270
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
                  Caption         =   "*.exe"
                  Enabled         =   0   'False
                  BackStyle       =   0
                  BackColor       =   -2147483639
                  ForeColor       =   0
               End
               Begin FVUnicodeControl.FVistaUniLabel lblLoaiDuoi 
                  Height          =   255
                  Left            =   120
                  TabIndex        =   160
                  Top             =   120
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   450
                  BackStyle       =   0
                  Caption         =   "Loa5i D9uo6i File Ca62n Ti2m:"
                  Enabled         =   0   'False
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
            End
            Begin FVUnicodeControl.FVistaUniButton cmdApplyNangCao 
               Height          =   255
               Left            =   5640
               TabIndex        =   156
               Top             =   3120
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   450
               BackColor       =   8421631
               ButtonShape     =   3
               ButtonStyle     =   1
               Caption         =   "A1p Du5ng"
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
            Begin FVUnicodeControl.FVistaUniOption optSearchNormal 
               Height          =   195
               Left            =   360
               TabIndex        =   155
               Top             =   600
               Width           =   2040
               _ExtentX        =   3598
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
               Value           =   -1  'True
               Caption         =   "Ti2m Kie61m To61c D9o65 Nhanh."
               BackStyle       =   0
               ForeColor       =   0
            End
            Begin FVUnicodeControl.FVistaUniOption optCungTenFolder 
               Height          =   195
               Left            =   360
               TabIndex        =   154
               Top             =   960
               Width           =   4095
               _ExtentX        =   7223
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
               Caption         =   "Ti2m Ca1c Ta65p Tin Co1 Te6n Tru2ng Vo71i Thu7 Mu5c Chu71a No1."
               BackStyle       =   0
               ForeColor       =   0
            End
            Begin FVUnicodeControl.FVistaUniButton FVistaUniButton16 
               Height          =   375
               Left            =   3000
               TabIndex        =   183
               Top             =   1200
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   661
               BackColor       =   14737632
               ButtonShape     =   3
               ButtonStyle     =   9
               Caption         =   "?"
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
         End
         Begin VB.Timer tmrFinding 
            Enabled         =   0   'False
            Interval        =   100
            Left            =   8160
            Top             =   120
         End
         Begin FVUnicodeControl.FVistaUniLabel lblStatus 
            Height          =   255
            Left            =   120
            TabIndex        =   151
            Top             =   4200
            Width           =   6855
            _ExtentX        =   12091
            _ExtentY        =   450
            BackColor       =   -2147483629
            Caption         =   ""
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
         Begin OneClick.McListBox ListFind1 
            Height          =   3375
            Left            =   120
            TabIndex        =   139
            Top             =   600
            Width           =   7935
            _ExtentX        =   13996
            _ExtentY        =   5953
            Picture         =   "frmMain.frx":1AAF2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   131072
            IconFocus       =   0   'False
            RowHeight       =   15
         End
         Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel3 
            Height          =   255
            Left            =   3120
            TabIndex        =   140
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            Alignment       =   2
            BackStyle       =   0
            Caption         =   "D9i5a Chi3 Ti2m:"
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
         Begin FVUnicodeControl.FVistaUniLabel lblKeyWord 
            Height          =   375
            Left            =   120
            TabIndex        =   141
            Top             =   240
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   661
            Caption         =   "Tu72 Kho1a:"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   0
         End
         Begin FVUnicodeControl.FVistaUniButton cmdSearch 
            Height          =   255
            Left            =   7200
            TabIndex        =   142
            Top             =   240
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   450
            BackColor       =   -2147483646
            ButtonShape     =   3
            ButtonStyle     =   1
            Caption         =   "Ti2m"
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
         Begin FVUnicodeControl.FVistaUniTextbox txtPath2Find 
            Height          =   270
            Left            =   4200
            TabIndex        =   143
            Top             =   240
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   476
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   0
            Text            =   "C:\"
            Locked          =   -1  'True
            BorderStyle     =   2
            BorderLine      =   11709605
         End
         Begin FVUnicodeControl.FVistaUniTextbox txtKeyWord 
            Height          =   270
            Left            =   960
            TabIndex        =   144
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   476
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   0
            Text            =   "*.*"
            BorderStyle     =   2
            BorderLine      =   11709605
         End
         Begin FVUnicodeControl.FVistaUniButton cmdSearchCaoCap 
            Height          =   375
            Left            =   8160
            TabIndex        =   152
            Top             =   600
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            BackColor       =   255
            ButtonShape     =   3
            ButtonStyle     =   1
            Caption         =   "Na6ng Cao"
            CheckBox        =   -1  'True
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
         Begin FVUnicodeControl.FVistaUniButton cmdDeleteFile 
            Height          =   255
            Left            =   8160
            TabIndex        =   161
            Top             =   1320
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
            BackColor       =   -2147483646
            ButtonShape     =   3
            ButtonStyle     =   1
            Caption         =   "Xo1a File D9a4 Cho5n"
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
         Begin FVUnicodeControl.FVistaUniButton cmdOpenFolderCha 
            Height          =   255
            Left            =   8160
            TabIndex        =   162
            Top             =   2040
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
            BackColor       =   -2147483646
            ButtonShape     =   3
            ButtonStyle     =   1
            Caption         =   "Mo73 Thu7 Mu5c Cha"
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
         Begin FVUnicodeControl.FVistaUniButton cmdDeleteAll 
            Height          =   255
            Left            =   8160
            TabIndex        =   163
            Top             =   1680
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
            BackColor       =   -2147483646
            ButtonShape     =   3
            ButtonStyle     =   1
            Caption         =   "Xo1a He61t"
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
      End
      Begin FVUnicodeControl.FVistaUniLabel lblPath 
         Height          =   255
         Left            =   -46540
         TabIndex        =   145
         Top             =   4680
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   450
         AutoUnicode     =   0   'False
         BackColor       =   -2147483629
         Caption         =   ""
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
      Begin FVUnicodeControl.FVistaUniCheckbox chkShowSystem 
         Height          =   195
         Left            =   -39700
         TabIndex        =   146
         Top             =   720
         Width           =   2190
         _ExtentX        =   3863
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
         Value           =   -1  'True
         Caption         =   "Hie65n File, Folder He65 Tho61ng"
         ForeColor       =   0
      End
      Begin FVUnicodeControl.FVistaUniCheckbox chkShowHidden 
         Height          =   195
         Left            =   -39700
         TabIndex        =   147
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
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
         Value           =   -1  'True
         Caption         =   "Hie65n File, Folder A63n"
         ForeColor       =   0
      End
      Begin OneClick.McListBox FileDir 
         Height          =   4095
         Left            =   -43660
         TabIndex        =   148
         Top             =   480
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   7223
         Picture         =   "frmMain.frx":1AB0E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   131072
         IconFocus       =   0   'False
         RowHeight       =   18
         BackGradient    =   3
         BackGradientCol =   -2147483629
         ShowIcon        =   -1  'True
         AutoHideScrollBars=   -1  'True
         Mode            =   3
         Path            =   "X:\"
         ShowSystemFiles =   -1  'True
         ShowHiddenFiles =   -1  'True
      End
      Begin OneClick.McListBox FolderDir 
         Height          =   4095
         Left            =   -46540
         TabIndex        =   149
         Top             =   480
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   7223
         Picture         =   "frmMain.frx":1AB2A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   131072
         IconFocus       =   0   'False
         RowHeight       =   18
         BackGradient    =   1
         BackGradientCol =   -2147483629
         ShowIcon        =   -1  'True
         AutoHideScrollBars=   -1  'True
         Mode            =   4
         Path            =   "D:\zVBLT\QuickFix\Q\"
         ShowSystemFiles =   -1  'True
         ShowHiddenFiles =   -1  'True
      End
      Begin FVUnicodeControl.FVistaUniButton FVistaUniButton11 
         Height          =   495
         Left            =   -8545
         TabIndex        =   165
         Top             =   1680
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   873
         BackColor       =   -2147483635
         ButtonShape     =   3
         Caption         =   "Ca2i D9a85t Registry To61t Nha61t"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PicAlign        =   3
         PicOpacity      =   0
      End
   End
   Begin VB.Timer tmrStart 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6600
      Top             =   120
   End
   Begin FVUnicodeControl.FVistaUniLabel lbllhelp 
      Height          =   255
      Left            =   1560
      TabIndex        =   93
      Top             =   1080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      Caption         =   "Hu7o71ng Da64n"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin FVUnicodeControl.FVistaUniLabel lblTacGia 
      Height          =   255
      Left            =   9240
      TabIndex        =   96
      Top             =   1080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      Alignment       =   1
      Caption         =   "Tho6ng Tin"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin FVUnicodeControl.FVistaUniLabel lblSetting 
      Height          =   255
      Left            =   0
      TabIndex        =   105
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      Caption         =   "Ca2i D9a85t Chung"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
   End
   Begin VB.Image Image1 
      Height          =   1260
      Left            =   3120
      Picture         =   "frmMain.frx":1AB46
      Top             =   120
      Width           =   3000
   End
   Begin VB.Image cmdCaiDatChung 
      Height          =   720
      Left            =   360
      MousePointer    =   2  'Cross
      Picture         =   "frmMain.frx":22442
      Top             =   240
      Width           =   720
   End
   Begin VB.Image cmdTacGia 
      Height          =   720
      Left            =   9360
      MousePointer    =   14  'Arrow and Question
      Picture         =   "frmMain.frx":27C24
      Top             =   240
      Width           =   720
   End
   Begin VB.Image cmdUpdate 
      Height          =   720
      Left            =   8040
      MousePointer    =   2  'Cross
      Picture         =   "frmMain.frx":2D406
      Top             =   240
      Width           =   720
   End
   Begin VB.Image cmdHelp 
      Height          =   720
      Left            =   1800
      MousePointer    =   2  'Cross
      Picture         =   "frmMain.frx":32BE8
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MAX_COMPUTERNAME_LENGTH As Long = 31
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal lpszConnectionName As String, ByVal dwNameLen As Integer, ByVal dwReserved As Long) As Long
Private Declare Function DefWindowProcW Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_SETTEXT As Long = &HC
 

Dim sStopSearch As Boolean


Dim sDangTim
Private Enum KieuSearch
    SearchBinhThuong = 0
    FileCungTenVoiThuMucChuaNo = 1
    SearchByIcon = 2
End Enum




'//////////////// Search File /////////////////////
Private Const vbDot = 46
Private Const MAXDWORD As Long = &HFFFFFFFF
Private Const MAX_PATH As Long = 260
Private Const INVALID_HANDLE_VALUE = -1
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10

Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type

Private Type FILE_PARAMS
   bRecurse As Boolean
   sFileRoot As String
   sFileNameExt As String
   sResult As String
   sMatches As String
   Count As Long
End Type

Private Declare Function FindClose Lib "kernel32" _
  (ByVal hFindFile As Long) As Long
   
Private Declare Function FindFirstFile Lib "kernel32" _
   Alias "FindFirstFileA" _
  (ByVal lpFileName As String, _
   lpFindFileData As WIN32_FIND_DATA) As Long
   
Private Declare Function FindNextFile Lib "kernel32" _
   Alias "FindNextFileA" _
  (ByVal hFindFile As Long, _
   lpFindFileData As WIN32_FIND_DATA) As Long

'/////////////////////////////////////////////////////////








Dim sConnType As String * 255
Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type
Private memInfo As MEMORYSTATUS
Dim memoryInfo As MEMORYSTATUS
Dim lastpcent As Single, lastTot As Long
Private Declare Sub GlobalMemoryStatus Lib "kernel32" _
   (lpBuffer As MEMORYSTATUS)



Dim sUnload As Boolean
Dim Inde As Integer
Dim sMoRong As Boolean



Private Sub AutorunHelp_MouseEnter()
lblHelp.Caption = "Kho6ng cho cha5y Autorun.inf tu72 ca1c o63 USB, vie65c na2y se4 giu1p ba5n tra1nh kho3i bi5 nhie64m Virus tu72 USB."
End Sub





Private Sub chkShowHidden_Click()
If chkShowHidden.Value = True Then
    FileDir.ShowHiddenFiles = True
    FolderDir.ShowHiddenFiles = True
Else
    FileDir.ShowHiddenFiles = False
    FolderDir.ShowHiddenFiles = False
End If
FileDir.Refresh
FolderDir.Refresh
End Sub

Private Sub chkShowSystem_Click()
If chkShowSystem.Value = True Then
    FolderDir.ShowSystemFiles = True
    FileDir.ShowSystemFiles = True
Else
    FolderDir.ShowSystemFiles = False
    FileDir.ShowSystemFiles = False
End If
FolderDir.Refresh
FileDir.Refresh
End Sub

Private Sub ClockHelp_MouseEnter()
lblHelp.Caption = "A63n D9o62ng Ho62 Va2 Ca1c Icon Tre6n Thanh Taskbar: Khi chu71c na8ng na2y bi5 kho1a, D9o62ng Ho62 Va2 Ca1c Icon Tre6n Thanh Taskbar se4 bie61n ma61t. D9a1nh da61u d9e63 kho1a chu71c na8ng na2y, bo3 d9a1nh da61u d9e63 mo73 chu71c na8ng na2y. Nha61n nu1t thu75c hie65n d9e63 thay d9o63 co1 hie65u lu75c."

End Sub

Private Sub cmdAdd_Click()
If txtFileName.Text = "" Then Exit Sub
Dim Stt As Integer
Dim sOK As Boolean
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "DisallowRun", 1
sOK = False
Dim i
For i = 0 To 100

If GetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun", i) = "" And sOK = False Then
sOK = True
SaveString HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun", i, txtFileName.Text
ListCam.AddItem txtFileName.Text
End If
Next i


txtFileName.Text = ""
End Sub

Private Sub cmdAdd_MouseEnter()
lblHelp.Caption = "Ne61u ba5n muo61n nga8n kho6ng cho cha5y File na2o d9o1 (VD: File Virus, ca1c File bi5 lo64i, v..v..). Chi3 ca62n go4 te6n File muo61n nga8n va2o khung va2 nha61n nu1t na2y, ca1c File d9a4 va2 d9ang d9u7o75c nga8n se4 hie63n thi5 o73 khung be6n ca5nh. Ba61t cu71 File na2o co1 te6n la2 mo6t trong so61 ca1c te6n o73 danh sa1ch be6n d9e62u kho6ng the63 cha5y. Nha61n nu1t Thu71c Hie65n D9e63 Ba81t D9a62u Nga8n. Lu7u Y1 !: File Va64n D9u7o75c Nga8n Khi Chu7o7ng Tri2nh Kho6ng Hoa5t D9o65ng."
End Sub

Private Sub cmdAddFile_Click()
If List1.ListCount < 10 Then
Dim sFileName
sFileName = DiaLog1.ShowOpen("All File (*.*)", , "C:\", "Add File")

If sFileName <> "" Then List1.AddItem sFileName


Else
UniMsgBox ChrW$(&H42) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H1EC9) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&HF3) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H68) & ChrW$(&H1EC3) & ChrW$(&H20) & ChrW$(&H78) & ChrW$(&HF3) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H31) & ChrW$(&H30) & ChrW$(&H20) & ChrW$(&H46) & ChrW$(&H69) & ChrW$(&H6C) & ChrW$(&H65) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H72) & ChrW$(&H6F) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H6D) & ChrW$(&H1ED9) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H6C) & ChrW$(&H1EA7) & ChrW$(&H6E) & ChrW$(&H21), vbOKOnly, "Thông báo", Me.hwnd
End If

End Sub

Private Sub cmdAddFile_MouseEnter()
lblHelp.Caption = "The6m File Va2o Danh Sa1ch Xo1a"
End Sub

Private Sub cmdApplyNangCao_Click()
'3015
If optCungTenFolder.Value = True Then
lblKeyWord.Caption = "[Ti2m File Cu2ng Te6n Vo71i Thu7 Mu5c Chu71a No1]"
lblKeyWord.Width = 3135
ElseIf optSearchNormal.Value = True Then
lblKeyWord.Caption = "Tu72 Kho1a:"
lblKeyWord.Width = 735
ElseIf optSearchByIcon.Value = True Then
lblKeyWord.Caption = "[Ti2m File Theo Icon]"
lblKeyWord.Width = 3135
End If

fmNangCao.Visible = False
cmdSearchCaoCap.Value = False
txtKeyWord.Enabled = True
txtPath2Find.Enabled = True
cmdSearch.Enabled = True
End Sub

Private Sub cmdCaiDatChung_Click()
If fmCaiDatChung.Visible = False Then
fmCaiDatChung.Visible = True
FVistaUniTabStrip1.Visible = False
GetComputerInfo
Else
fmCaiDatChung.Visible = False
FVistaUniTabStrip1.Visible = True
End If

End Sub

Private Sub cmdCaiDatChung_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdCaiDatChung.BorderStyle = 1
End Sub

Private Sub cmdCaiDatChung_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdCaiDatChung.BorderStyle = 0
End Sub

Private Sub cmdCancel_Click()
List1.Clear
DeleteValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\Session Manager", "PendingFileRenameOperations"
UniMsgBox ChrW$(&H4B) & ChrW$(&H68) & ChrW$(&HF4) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H58) & ChrW$(&HF3) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H46) & ChrW$(&H69) & ChrW$(&H6C) & ChrW$(&H65) & ChrW$(&H20) & ChrW$(&H4E) & ChrW$(&H1EEF) & ChrW$(&H61) & ChrW$(&H2E), vbOKOnly, "Thông Báo", Me.hwnd
End Sub

Private Sub cmdCancel_MouseEnter()
lblHelp.Caption = "Kho6ng Xo1a File Nu74a."
End Sub

Private Sub cmdChangeInfo_Click()
Dim sNoiDungGhi As String
If FileExists("C:\WINDOWS\system32\Oemlogo.bmp") = True Then
DeleteFile "C:\WINDOWS\system32\Oemlogo.bmp"
End If
SavePicture PicOemLogo.Image, "C:\WINDOWS\system32\Oemlogo.bmp"
sNoiDungGhi = "[General]" & vbCrLf _
& "Manufacturer=1Click - Quang Trung Software" & vbCrLf _
& "Model=Nhanh Chong - Don Gian" & vbCrLf _
& "[Support Information]" & vbCrLf _
& "line1=" & LayDong(1, frmMain.txtOemInfo.Text) & vbCrLf _
& "line2=" & LayDong(2, frmMain.txtOemInfo.Text) & vbCrLf _
& "line3=" & LayDong(3, frmMain.txtOemInfo.Text) & vbCrLf _
& "line4=" & LayDong(4, frmMain.txtOemInfo.Text) & vbCrLf _
& "line5=" & LayDong(5, frmMain.txtOemInfo.Text) & vbCrLf _
& "line6=" & LayDong(6, frmMain.txtOemInfo.Text) & vbCrLf _
& "line7=" & LayDong(7, frmMain.txtOemInfo.Text) & vbCrLf _
& "line8=" & LayDong(8, frmMain.txtOemInfo.Text) & vbCrLf _
& "line9=" & LayDong(9, frmMain.txtOemInfo.Text) & vbCrLf _
& "line10=" & LayDong(10, frmMain.txtOemInfo.Text)
CreateTextFile "C:\WINDOWS\system32\OEMINFO.INI", sNoiDungGhi
UniMsgBox ChrW(272) & ChrW(227) & " " & ChrW(273) & ChrW(7893) & "i th" & ChrW(244) & "ng tin th" & ChrW(224) & "nh c" & ChrW(244) & "ng.", vbOKOnly, "Thông Báo", Me.hwnd
Dim dblReturn As Double
dblReturn = Shell("rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl @1", 5)
End Sub

Private Sub cmdChangePicture_Click()
Dim sBitMapFile
sBitMapFile = DiaLog1.ShowOpen("Image Files" + Chr(0) + "*.bmp" + Chr(0), , "C:\", "Image File")
If sBitMapFile <> "" Then
PicOemLogo.Picture = LoadPicture(sBitMapFile)
End If
End Sub

Private Sub cmdCheckVR_Click()
Dim CoVirus As Boolean
CoVirus = False
'C:\WINDOWS\system32\drivers\klif.sys
If FileExists("C:\2fiy.bat") = True Then
UniMsgBox ChrW$(&H4D) & ChrW$(&HE1) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&HED) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H1EE7) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1ECB) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H69) & ChrW$(&H1EC5) & ChrW$(&H6D) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H3A) & ChrW$(&H20) & " 2fiy.bat" _
& ChrW$(&H2E) & ChrW$(&H48) & ChrW$(&HE3) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&HF9) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H1EE9) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H103) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H61) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1EC3) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&HE0) & ChrW$(&H79) & ChrW$(&H2E), vbOKOnly, "Có Virus !!!", Me.hwnd
CoVirus = True
End If



If FileExists("C:\WINDOWS\system32\win1ogon.exe") = True Then
UniMsgBox ChrW$(&H4D) & ChrW$(&HE1) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&HED) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H1EE7) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1ECB) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H69) & ChrW$(&H1EC5) & ChrW$(&H6D) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H3A) & ChrW$(&H20) & " win1ogon.exe" _
& ChrW$(&H2E) & ChrW$(&H48) & ChrW$(&HE3) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&HF9) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H1EE9) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H103) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H61) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1EC3) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&HE0) & ChrW$(&H79) & ChrW$(&H2E), vbOKOnly, "Có Virus !!!", Me.hwnd
CoVirus = True
End If


If FileExists("C:\WINDOWS\system32\SCVVHSOT.exe") = True Then
UniMsgBox ChrW$(&H4D) & ChrW$(&HE1) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&HED) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H1EE7) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1ECB) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H69) & ChrW$(&H1EC5) & ChrW$(&H6D) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H3A) & ChrW$(&H20) & " SCVVHSOT.exe" _
& ChrW$(&H2E) & ChrW$(&H48) & ChrW$(&HE3) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&HF9) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H1EE9) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H103) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H61) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1EC3) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&HE0) & ChrW$(&H79) & ChrW$(&H2E), vbOKOnly, "Có Virus !!!", Me.hwnd
CoVirus = True
End If


If FileExists("C:\WINDOWS\system32\Explorer.sm1") = True Then
UniMsgBox ChrW$(&H4D) & ChrW$(&HE1) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&HED) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H1EE7) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1ECB) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H69) & ChrW$(&H1EC5) & ChrW$(&H6D) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H3A) & ChrW$(&H20) & " sxs.exe" _
& ChrW$(&H2E) & ChrW$(&H48) & ChrW$(&HE3) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&HF9) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H1EE9) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H103) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H61) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1EC3) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&HE0) & ChrW$(&H79) & ChrW$(&H2E), vbOKOnly, "Có Virus !!!", Me.hwnd
CoVirus = True
End If

If FileExists("C:\WINDOWS\system32\Explorer.sm1") = True Then
UniMsgBox ChrW$(&H4D) & ChrW$(&HE1) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&HED) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H1EE7) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1ECB) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H69) & ChrW$(&H1EC5) & ChrW$(&H6D) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H3A) & ChrW$(&H20) & " sxs.exe" _
& ChrW$(&H2E) & ChrW$(&H48) & ChrW$(&HE3) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&HF9) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H1EE9) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H103) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H61) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1EC3) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&HE0) & ChrW$(&H79) & ChrW$(&H2E), vbOKOnly, "Có Virus !!!", Me.hwnd
CoVirus = True
End If

If FileExists("C:\WINDOWS\megabyte.exe") = True Then
UniMsgBox ChrW$(&H4D) & ChrW$(&HE1) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&HED) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H1EE7) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1ECB) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H69) & ChrW$(&H1EC5) & ChrW$(&H6D) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H3A) & ChrW$(&H20) & " Megabyte.exe" _
& ChrW$(&H2E) & ChrW$(&H48) & ChrW$(&HE3) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&HF9) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H1EE9) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H103) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H61) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1EC3) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&HE0) & ChrW$(&H79) & ChrW$(&H2E), vbOKOnly, "Có Virus !!!", Me.hwnd
CoVirus = True

End If

If FileExists("C:\WINDOWS\System32\Sys\HayTiepTuc.exe") = True Then
UniMsgBox ChrW$(&H4D) & ChrW$(&HE1) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&HED) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H1EE7) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1ECB) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H69) & ChrW$(&H1EC5) & ChrW$(&H6D) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H3A) & ChrW$(&H20) & " TiepTuc.exe" _
& ChrW$(&H2E) & ChrW$(&H48) & ChrW$(&HE3) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&HF9) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H1EE9) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H103) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H61) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1EC3) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&HE0) & ChrW$(&H79) & ChrW$(&H2E), vbOKOnly, "Có Virus !!!", Me.hwnd
CoVirus = True

End If

If FileExists("C:\RESTORE\S-1-5-21-1482476501-1644491937-682003330-1013\Taquito.exe") = True Then
UniMsgBox ChrW$(&H4D) & ChrW$(&HE1) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&HED) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H1EE7) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1ECB) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H69) & ChrW$(&H1EC5) & ChrW$(&H6D) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H3A) & ChrW$(&H20) & " Taquito.exe" _
& ChrW$(&H2E) & ChrW$(&H48) & ChrW$(&HE3) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&HF9) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H1EE9) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H103) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H61) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1EC3) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&HE0) & ChrW$(&H79) & ChrW$(&H2E), vbOKOnly, "Có Virus !!!", Me.hwnd
CoVirus = True

End If

If FileExists("C:\Folder.exe") = True Then
UniMsgBox ChrW$(&H4D) & ChrW$(&HE1) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&HED) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H1EE7) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1ECB) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H69) & ChrW$(&H1EC5) & ChrW$(&H6D) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H3A) & ChrW$(&H20) & " EV-SHUTTLE.exe" _
& ChrW$(&H2E) & ChrW$(&H48) & ChrW$(&HE3) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&HF9) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H1EE9) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H103) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H61) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1EC3) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&HE0) & ChrW$(&H79) & ChrW$(&H2E), vbOKOnly, "Có Virus !!!", Me.hwnd
CoVirus = True

End If


If FileExists("C:\WINDOWS\taskmsg.exe") = True Then
UniMsgBox ChrW$(&H4D) & ChrW$(&HE1) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&HED) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H1EE7) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1ECB) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H69) & ChrW$(&H1EC5) & ChrW$(&H6D) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H3A) & ChrW$(&H20) & " taskmsg.exe" _
& ChrW$(&H2E) & ChrW$(&H48) & ChrW$(&HE3) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&HF9) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H1EE9) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H103) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H61) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1EC3) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&HE0) & ChrW$(&H79) & ChrW$(&H2E), vbOKOnly, "Có Virus !!!", Me.hwnd
CoVirus = True

End If

If FileExists("C:\WINDOWS\system32\scvhosti.exe") = True Then
UniMsgBox ChrW$(&H4D) & ChrW$(&HE1) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&HED) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H1EE7) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1ECB) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H69) & ChrW$(&H1EC5) & ChrW$(&H6D) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H3A) & ChrW$(&H20) & " scvhosti.exe" _
& ChrW$(&H2E) & ChrW$(&H48) & ChrW$(&HE3) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&HF9) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H1EE9) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H103) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H61) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1EC3) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&HE0) & ChrW$(&H79) & ChrW$(&H2E), vbOKOnly, "Có Virus !!!", Me.hwnd
CoVirus = True

End If

If FileExists("C:\Program Files\PCPrivacyCleaner\pcpc.exe") = True Then
UniMsgBox ChrW$(&H4D) & ChrW$(&HE1) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&HED) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H1EE7) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1ECB) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H69) & ChrW$(&H1EC5) & ChrW$(&H6D) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H3A) & ChrW$(&H20) & " pcpc.exe" _
& ChrW$(&H2E) & ChrW$(&H48) & ChrW$(&HE3) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&HF9) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H1EE9) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H103) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H61) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1EC3) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&HE0) & ChrW$(&H79) & ChrW$(&H2E), vbOKOnly, "Có Virus !!!", Me.hwnd
CoVirus = True

End If

If FileExists("C:\WINDOWS\System32\logon.exe") = True Then
UniMsgBox ChrW$(&H4D) & ChrW$(&HE1) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&HED) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H1EE7) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1ECB) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H69) & ChrW$(&H1EC5) & ChrW$(&H6D) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H3A) & ChrW$(&H20) & " algs.exe" _
& ChrW$(&H2E) & ChrW$(&H48) & ChrW$(&HE3) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&HF9) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H1EE9) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H103) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H61) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1EC3) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&HE0) & ChrW$(&H79) & ChrW$(&H2E), vbOKOnly, "Có Virus !!!", Me.hwnd
CoVirus = True

End If

If FileExists("C:\Program Files\AntiMalwareGuard\amg.exe") = True Then
UniMsgBox ChrW$(&H4D) & ChrW$(&HE1) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&HED) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H1EE7) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1ECB) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H69) & ChrW$(&H1EC5) & ChrW$(&H6D) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H3A) & ChrW$(&H20) & " amg.exe" _
& ChrW$(&H2E) & ChrW$(&H48) & ChrW$(&HE3) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&HF9) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H1EE9) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H103) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H61) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1EC3) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&HE0) & ChrW$(&H79) & ChrW$(&H2E), vbOKOnly, "Có Virus !!!", Me.hwnd
 CoVirus = True

End If

If FileExists("C:\WINDOWS\system32\IEXPLORER.exe") = True Then
UniMsgBox ChrW$(&H4D) & ChrW$(&HE1) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&HED) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H1EE7) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1ECB) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H69) & ChrW$(&H1EC5) & ChrW$(&H6D) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H3A) & ChrW$(&H20) & " IEXPLORER.exe" _
& ChrW$(&H2E) & ChrW$(&H48) & ChrW$(&HE3) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&HF9) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H1EE9) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H103) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H61) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1EC3) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&HE0) & ChrW$(&H79) & ChrW$(&H2E), vbOKOnly, "Có Virus !!!", Me.hwnd
CoVirus = True

End If

If FileExists("C:\WINDOWS\help\B7C8A6484EE3.exe") = True Then
UniMsgBox ChrW$(&H4D) & ChrW$(&HE1) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&HED) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H1EE7) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1ECB) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H69) & ChrW$(&H1EC5) & ChrW$(&H6D) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H3A) & ChrW$(&H20) & " Shell.exe" _
& ChrW$(&H2E) & ChrW$(&H48) & ChrW$(&HE3) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&HF9) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H1EE9) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H103) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H61) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1EC3) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&HE0) & ChrW$(&H79) & ChrW$(&H2E), vbOKOnly, "Có Virus !!!", Me.hwnd
CoVirus = True

End If


If FileExists("C:\WINDOWS\System32\system.exe") = True Then
UniMsgBox ChrW$(&H4D) & ChrW$(&HE1) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&HED) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H1EE7) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1ECB) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H69) & ChrW$(&H1EC5) & ChrW$(&H6D) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H3A) & ChrW$(&H20) & " Forever.exe" _
& ChrW$(&H2E) & ChrW$(&H48) & ChrW$(&HE3) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&HF9) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H1EE9) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H103) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H61) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1EC3) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&HE0) & ChrW$(&H79) & ChrW$(&H2E), vbOKOnly, "Có Virus !!!", Me.hwnd
CoVirus = True

End If


If FileExists("C:\lbb.com") = True Then
UniMsgBox ChrW$(&H4D) & ChrW$(&HE1) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&HED) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H1EE7) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1ECB) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H69) & ChrW$(&H1EC5) & ChrW$(&H6D) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H3A) & ChrW$(&H20) & " KvoSoft" _
& ChrW$(&H2E) & ChrW$(&H48) & ChrW$(&HE3) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&HF9) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H1EE9) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H103) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H61) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1EC3) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&HE0) & ChrW$(&H79) & ChrW$(&H2E), vbOKOnly, "Có Virus !!!", Me.hwnd
CoVirus = True

End If



If FileExists("C:\zPharaoh.exe") = True Then
UniMsgBox ChrW$(&H4D) & ChrW$(&HE1) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&HED) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H1EE7) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1ECB) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H69) & ChrW$(&H1EC5) & ChrW$(&H6D) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H3A) & ChrW$(&H20) & " zPharaoh.exe" _
& ChrW$(&H2E) & ChrW$(&H48) & ChrW$(&HE3) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&HF9) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H1EE9) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H103) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H61) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1EC3) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&HE0) & ChrW$(&H79) & ChrW$(&H2E), vbOKOnly, "Có Virus !!!", Me.hwnd
CoVirus = True

End If

If FileExists("C:\WINDOWS\phimnguoilon.exe") = True Then
UniMsgBox ChrW$(&H4D) & ChrW$(&HE1) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&HED) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H1EE7) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1ECB) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H69) & ChrW$(&H1EC5) & ChrW$(&H6D) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H3A) & ChrW$(&H20) & " Phimhot.exe" _
& ChrW$(&H2E) & ChrW$(&H48) & ChrW$(&HE3) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&HF9) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H1EE9) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H103) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H61) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1EC3) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&HE0) & ChrW$(&H79) & ChrW$(&H2E), vbOKOnly, "Có Virus !!!", Me.hwnd
CoVirus = True

End If


If FileExists("C:\WINDOWS\Mixa.exe") = True Then
UniMsgBox ChrW$(&H4D) & ChrW$(&HE1) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&HED) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H1EE7) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1ECB) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H69) & ChrW$(&H1EC5) & ChrW$(&H6D) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H3A) & ChrW$(&H20) & " Mixa_I.exe" _
& ChrW$(&H2E) & ChrW$(&H48) & ChrW$(&HE3) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&HF9) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H1EE9) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H103) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H61) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1EC3) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&HE0) & ChrW$(&H79) & ChrW$(&H2E), vbOKOnly, "Có Virus !!!", Me.hwnd
CoVirus = True

End If

If CoVirus = False Then
UniMsgBox ChrW(272) & ChrW(227) & " qu" & ChrW(233) & "t xong, kh" & ChrW(244) & "ng t" & ChrW(236) & "m th" & ChrW(7845) & "y Virus n" & ChrW(224) & "o.", vbOKOnly, "Thông Báo", Me.hwnd
Else
UniMsgBox ChrW$(&H110) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H71) & ChrW$(&H75) & ChrW$(&HE9) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H78) & ChrW$(&H6F) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H2C) & ChrW$(&H20) & ChrW$(&H68) & ChrW$(&HE3) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&HF9) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H1EE9) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H103) _
& ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H22) & ChrW$(&H44) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H2D) & ChrW$(&H20) & ChrW$(&H31) & ChrW$(&H43) & ChrW$(&H6C) & ChrW$(&H69) & ChrW$(&H63) & ChrW$(&H6B) & ChrW$(&H20) & ChrW$(&H6C) & ChrW$(&HE0) & ChrW$(&H20) & ChrW$(&H78) & ChrW$(&H6F) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H22) & ChrW$(&H20) & ChrW$(&H111) _
& ChrW$(&H1EC3) & ChrW$(&H20) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H2E), vbOKOnly, "Thông Báo", Me.hwnd
End If
End Sub

Private Sub cmdCheckVR_MouseEnter()
lblHelp.Caption = "Que1t Virus: Chu7o7ng tri2nh se4 kie63m tra xem ma1y ti1nh cu3a ba5n co1 nhie64m Virus hay kho6ng (Lu7u y1 chi3 nhu74ng Virus co1 trong co7 so73 du74 lie65u cu3a chu7o7ng tri2nh thi2 chu7o7ng tri2nh na2y mo71i pha1t hie65n d9u7o75c). Ne61u co1 thi2 chu7o7ng tri2nh se4 tho6ng ba1o cho ba5n bie61t trong lu1c que1t, ne61u kho6ng co1 thi2 chu7o7ng tri2nh se4 kho6ng tho6ng ba1o. Co1 the63 xa3y ra tru7o72ng ho75p ma1y co1 Virus nhu7ng chu7o7ng tri2nh kho6ng pha1t hie65n ra vi2 no1 kho6ng co1 trong co7 so73 du74 lie65u. Ba5n co1 the63 lie6n he65 vo71i ta1c gia3 va2 gu73i ma64u Virus le6n, Virus d9o1 se4 co1 trong co7 so73 du74 lue65u cu73a chu7o7ng tri2nh trong phie6n ba3n sau."
End Sub

Private Sub cmdDel_Click()
Dim St As Integer
Dim sOK As Boolean

sOK = False
If ListCam.ListIndex <> -1 Then
For St = 0 To 100
On Error Resume Next
If GetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun", St) = ListCam.ItemText(ListCam.ListIndex) And sOK = False Then
sOK = True
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun", St
End If
Next St
ListCam.Clear
UpdateList
Else
UniMsgBox ChrW$(&H42) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H1B0) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H1ECD) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H46) & ChrW$(&H69) & ChrW$(&H6C) & ChrW$(&H65), vbOKOnly, "Thông Báo", Me.hwnd
End If
End Sub



Private Sub cmdDel_MouseEnter()
lblHelp.Caption = "Xo1a te6n File d9a4 cho5n ra kho3i danh sa1ch d9o62ng tho72i cho phe1p File d9u7o75c cha5y."
End Sub

Private Sub cmdDeleteAll_Click()
If Not ListFind1.ListIndex = -1 Then
If UniMsgBox("B" & ChrW(7841) & "n c" & ChrW(243) & " ch" & ChrW(7855) & "c ch" & ChrW(7855) & "n mu" & ChrW(7889) & "n x" & ChrW(243) & "a t" & ChrW(7845) & "t c" & ChrW(7843) & " c" & ChrW(225) & "c File trong danh s" & ChrW(225) & "ch kh" & ChrW(244) & "ng?", vbCritical + vbYesNo, "C" & ChrW(7849) & "n th" & ChrW(7853) & "n!!!", frmMain.hwnd) = vbYes Then
If UniMsgBox("Ch" & ChrW(432) & ChrW(417) & "ng tr" & ChrW(236) & "nh s" & ChrW(7869) & " x" & ChrW(243) & "a t" & ChrW(7845) & "t c" & ChrW(7843) & " c" & ChrW(225) & "c File c" & ChrW(243) & " trong danh s" & ChrW(225) & "ch, b" & ChrW(7841) & "n n" & ChrW(234) & "n xem k" & ChrW(7929) & " l" & ChrW(7841) & "i n" & ChrW(7871) & "u nh" & ChrW(432) & " danh s" & ChrW(225) & "ch qu" & ChrW(225) & " d" & ChrW(224) & "i. B" & ChrW(7841) & "n v" & ChrW(7851) & "n mu" & ChrW(7889) & "n x" & ChrW(243) & "a?", vbYesNo + vbExclamation, "!!!", frmMain.hwnd) = vbYes Then
Dim sPath
Dim i

For i = 0 To ListFind1.ListCount - 1
    sPath = ListFind1.List(i)
    DeleteFile sPath
Next i
ListFind1.Clear
UniMsgBox ChrW(272) & ChrW(227) & " X" & ChrW(243) & "a H" & ChrW(7871) & "t !", vbInformation + vbOKOnly, "!", frmMain.hwnd
End If
End If
End If
End Sub

Private Sub cmdDeleteAll_MouseEnter()
lblHelp.Caption = "Xo1a he61t ta61t ca3 ca1c File ti2m d9u7o75c. Ca63n tha65n vo72i chu71c na8ng na2y, tra1nh xo1a nha62m File he65 tho61ng."
End Sub

Private Sub cmdDeleteFile_Click()
On Error Resume Next
If Not ListFind1.ListIndex = -1 Then


If UniMsgBox("B" & ChrW(7841) & "n c" & ChrW(243) & " mu" & ChrW(7889) & "n x" & ChrW(243) & "a File n" & ChrW(224) & "y ra kh" & ChrW(244) & "ng?", vbYesNo, "?", frmMain.hwnd) = vbYes Then
Dim sPath
sPath = frmMain.lblStatus.Caption
DeleteFile sPath
ListFind1.Remove ListFind1.ListIndex
UniMsgBox ChrW(272) & ChrW(227) & " x" & ChrW(243) & "a xong.", vbOKOnly, "Thông Báo", frmMain.hwnd
End If

End If
End Sub

Private Sub cmdDeleteFile_MouseEnter()
lblHelp.Caption = "Xo1a File d9a4 cho5n trong danh sa1ch."
End Sub

Private Sub cmdDelFile_Click()
If List1.ListIndex <> -1 Then
List1.RemoveItem List1.ListIndex
End If
End Sub

Private Sub cmdDelFile_MouseEnter()
lblHelp.Caption = "Xo1a File Ra Kho3i Danh Sa1ch Xo1a."
End Sub

Private Sub cmdDiet_Click()
'2fiy.bat
If ListVR.List(ListVR.ListIndex) = "Mixa_I.exe" Then KillVirus.KillMixa
If ListVR.List(ListVR.ListIndex) = "Phimhot.exe" Then KillVirus.KillPhimHot
If ListVR.List(ListVR.ListIndex) = "Images.exe" Then KillVirus.KillImages
If ListVR.List(ListVR.ListIndex) = "zPharaoh.exe" Then KillVirus.KillzPharaoh
If ListVR.List(ListVR.ListIndex) = "Kvosoft.exe" Then KillVirus.KillKvoSoft
If ListVR.List(ListVR.ListIndex) = "Forever.exe" Then KillVirus.KillForever
If ListVR.List(ListVR.ListIndex) = "Shell.exe" Then KillVirus.KillShell
If ListVR.List(ListVR.ListIndex) = "Algs.exe" Then KillVirus.KillALGS
If ListVR.List(ListVR.ListIndex) = "Amg.exe" Then KillVirus.KillAmg
If ListVR.List(ListVR.ListIndex) = "IEXPLORER.EXE" Then KillVirus.KillIexplorer
If ListVR.List(ListVR.ListIndex) = "pcpc.exe" Then KillVirus.KillPCPC
If ListVR.List(ListVR.ListIndex) = "scvhosti.exe" Then KillVirus.KillScvhosti
If ListVR.List(ListVR.ListIndex) = "taskmsg.exe" Then KillVirus.KillTaskmsg
If ListVR.List(ListVR.ListIndex) = "EV-SHUTTLE.exe" Then KillVirus.KillEVShuttle
If ListVR.List(ListVR.ListIndex) = "Taquito.exe" Then KillVirus.KillTaquito
If ListVR.List(ListVR.ListIndex) = "TiepTuc.exe" Then KillVirus.KillTiepTuc
If ListVR.List(ListVR.ListIndex) = "Megabyte.exe" Then KillVirus.KillMegabyte
If ListVR.List(ListVR.ListIndex) = "sxs.exe" Then KillVirus.KillSxS
If ListVR.List(ListVR.ListIndex) = "SCVVHSOT.exe" Then KillVirus.KillSCVVHSOT
If ListVR.List(ListVR.ListIndex) = "win1ogon.exe" Then KillVirus.KillWin1ogon
If ListVR.List(ListVR.ListIndex) = "2fiy.bat" Then KillVirus.Kill2FIY

CleanReg
End Sub

Private Sub cmdDiet_MouseEnter()
lblHelp.Caption = "Ne61u ba5n bie61t (hoa85c nghi ngo72) ra82ng ma1y ti1nh cu73a ba5n d9ang bi5 nhie64m mo65t trong ca1c loa5i Virus co1 trong danh sa1ch tre6n, khi d9o1 ban5 chi3 ca62n nha61p va2o te6n Virus ma2 ba5n nghi ngo72 la2 ba5n bi5 nhie64m, chu7o7ng tri2nh se4 to61ng co63 no1 ra kho3i ma1y ti1nh cu73a ba5n ngay la65p tu71c, d9a3m ba3o sa5ch ta65n go61c."
End Sub

Private Sub cmdHelp_Click()
frmHelpMe.Show , Me
End Sub

Private Sub cmdHelp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdHelp.BorderStyle = 1
End Sub

Private Sub cmdHelp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdHelp.BorderStyle = 0
End Sub

Private Sub cmdKiemtra_Click()
On Error Resume Next
txtCheck.Text = "Log File Ki" & ChrW(7875) & "m Tra H" & ChrW(7879) & " Th" & ChrW(7889) & "ng - 1Click" & vbCrLf _
& vbCrLf _
& vbCrLf _
& "C" & ChrW(225) & "c Proccess " & ChrW(273) & "ang ch" & ChrW(7841) & "y:" & vbCrLf
Dim colitems
Dim objitem
Dim a
Set colitems = GetObject("winmgmts:\root\CIMV2").ExecQuery("SELECT * FROM Win32_Process")
   For Each objitem In colitems
      a = a & vbCrLf & "----------------------------------------------------------"
      a = a & vbCrLf & "Proccess: " & objitem.Caption
      a = a & vbCrLf & "File: " & objitem.ExecutablePath
      a = a & vbCrLf & "Handle: " & objitem.handle
      a = a & vbCrLf & "HandleCount: " & objitem.HandleCount
      a = a & vbCrLf & "ParentProcessId: " & objitem.ParentProcessId
      a = a & vbCrLf & "Priority: " & objitem.Priority
      a = a & vbCrLf & "ProcessId: " & objitem.ProcessId
      a = a & vbCrLf & "ThreadCount: " & objitem.ThreadCount
   Next
frmMain.txtCheck.Text = frmMain.txtCheck.Text & a
txtCheck.Text = txtCheck.Text & vbCrLf _
& vbCrLf _
& vbCrLf _
& "C" & ChrW(225) & "c kh" & ChrW(243) & "a ki" & ChrW(7875) & "m tra kh" & ChrW(7903) & "i " & ChrW(273) & ChrW(7897) & "ng:" & vbCrLf
GetKeyValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"
GetKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
GetKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
GetKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnceEx"
GetKeyValue "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\RunOnce"
GetKeyValue "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\RunServices"

txtCheck.Text = txtCheck.Text & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Ki" & ChrW(7875) & "m tra kh" & ChrW(243) & "a Winlogon:" & vbCrLf
GetKeyValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"


txtCheck.Text = txtCheck.Text & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "K" & ChrW(237) & "ch th" & ChrW(432) & ChrW(7899) & "c c" & ChrW(225) & "c File h" & ChrW(7879) & " th" & ChrW(7889) & "ng quan tr" & ChrW(7885) & "ng:" & vbCrLf
txtCheck.Text = txtCheck.Text & vbCrLf & "C:\NTDETECT.COM |" & FileLen("C:\NTDETECT.COM") & " Bytes"
txtCheck.Text = txtCheck.Text & vbCrLf & "C:\ntldr |" & FileLen("C:\ntldr") & " Bytes"
txtCheck.Text = txtCheck.Text & vbCrLf & "C:\AUTOEXEC.BAT |" & FileLen("C:\AUTOEXEC.BAT") & " Bytes"
txtCheck.Text = txtCheck.Text & vbCrLf & "C:\WINDOWS\regedit.exe |" & FileLen("C:\WINDOWS\regedit.exe") & " Bytes |" & " M" & ChrW(7863) & "c " & ChrW(273) & ChrW(7883) & "nh: 146432 Bytes"
txtCheck.Text = txtCheck.Text & vbCrLf & "C:\WINDOWS\NOTEPAD.exe |" & FileLen("C:\WINDOWS\NOTEPAD.exe") & " Bytes |" & " M" & ChrW(7863) & "c " & ChrW(273) & ChrW(7883) & "nh: 69120 Bytes"
txtCheck.Text = txtCheck.Text & vbCrLf & "C:\WINDOWS\explorer.exe |" & FileLen("C:\WINDOWS\explorer.exe") & " Bytes |" & " M" & ChrW(7863) & "c " & ChrW(273) & ChrW(7883) & "nh: 1033216 Bytes"
txtCheck.Text = txtCheck.Text & vbCrLf & "C:\WINDOWS\System32\winlogon.exe |" & FileLen("C:\WINDOWS\System32\winlogon.exe") & " Bytes |" & " M" & ChrW(7863) & "c " & ChrW(273) & ChrW(7883) & "nh: 502272 Bytes"
txtCheck.Text = txtCheck.Text & vbCrLf & "C:\WINDOWS\System32\services.exe |" & FileLen("C:\WINDOWS\System32\services.exe") & " Bytes |" & " M" & ChrW(7863) & "c " & ChrW(273) & ChrW(7883) & "nh: 108032 Bytes"
txtCheck.Text = txtCheck.Text & vbCrLf & "C:\WINDOWS\System32\cmd.exe |" & FileLen("C:\WINDOWS\System32\cmd.exe") & " Bytes |" & " M" & ChrW(7863) & "c " & ChrW(273) & ChrW(7883) & "nh: 388608 Bytes"
txtCheck.Text = txtCheck.Text & vbCrLf & "C:\WINDOWS\System32\svchost.exe |" & FileLen("C:\WINDOWS\System32\svchost.exe") & " Bytes |" & " M" & ChrW(7863) & "c " & ChrW(273) & ChrW(7883) & "nh: 14336 Bytes"
txtCheck.Text = txtCheck.Text & vbCrLf & "C:\WINDOWS\System32\userinit.exe |" & FileLen("C:\WINDOWS\System32\userinit.exe") & " Bytes |" & " M" & ChrW(7863) & "c " & ChrW(273) & ChrW(7883) & "nh: 24576 Bytes"
txtCheck.Text = txtCheck.Text & vbCrLf & "C:\WINDOWS\System32\mmc.exe |" & FileLen("C:\WINDOWS\System32\mmc.exe") & " Bytes |" & " M" & ChrW(7863) & "c " & ChrW(273) & ChrW(7883) & "nh: 815104 Bytes"
txtCheck.Text = txtCheck.Text & vbCrLf & "C:\WINDOWS\System32\shell32.dll |" & FileLen("C:\WINDOWS\System32\shell32.dll") & " Bytes |" & " M" & ChrW(7863) & "c " & ChrW(273) & ChrW(7883) & "nh: 8454656 Bytes"


txtCheck.Text = txtCheck.Text & vbCrLf & vbCrLf & vbCrLf & "File Autorun.inf trong " & ChrW(7893) & " C:\" & vbCrLf
txtCheck.Text = txtCheck.Text & "______________________" & vbCrLf
txtCheck.Text = txtCheck.Text & ReadTextFile("C:\autorun.inf") & vbCrLf
txtCheck.Text = txtCheck.Text & "______________________" & vbCrLf
txtCheck.Text = txtCheck.Text & vbCrLf & vbCrLf & vbCrLf & "File Autorun.inf trong " & ChrW(7893) & " D:\" & vbCrLf
txtCheck.Text = txtCheck.Text & "______________________" & vbCrLf
txtCheck.Text = txtCheck.Text & ReadTextFile("D:\autorun.inf") & vbCrLf
txtCheck.Text = txtCheck.Text & "______________________" & vbCrLf
txtCheck.Text = txtCheck.Text & vbCrLf & vbCrLf & vbCrLf & "File Autorun.inf trong " & ChrW(7893) & " E:\" & vbCrLf
txtCheck.Text = txtCheck.Text & "______________________" & vbCrLf
txtCheck.Text = txtCheck.Text & ReadTextFile("E:\autorun.inf") & vbCrLf
txtCheck.Text = txtCheck.Text & "______________________" & vbCrLf
txtCheck.Text = txtCheck.Text & vbCrLf & vbCrLf & vbCrLf & "File Autorun.inf trong " & ChrW(7893) & " F:\" & vbCrLf
txtCheck.Text = txtCheck.Text & "______________________" & vbCrLf
txtCheck.Text = txtCheck.Text & ReadTextFile("F:\autorun.inf") & vbCrLf
txtCheck.Text = txtCheck.Text & "______________________" & vbCrLf

If UniMsgBox("M" & ChrW(225) & "y t" & ChrW(237) & "nh c" & ChrW(7911) & "a b" & ChrW(7841) & "n " & ChrW(273) & ChrW(227) & " " & ChrW(273) & ChrW(432) & ChrW(7907) & "c ki" & ChrW(7875) & "m tra, h" & ChrW(227) & "y l" & ChrW(432) & "u log File n" & ChrW(224) & "y l" & ChrW(7841) & "i v" & ChrW(224) & " g" & ChrW(7917) & "i n" & ChrW(243) & " " & ChrW(273) & ChrW(7871) & "n nh" & ChrW(7919) & "ng ng" & ChrW(432) & ChrW(7901) & "i c" & ChrW(243) & " chuy" & ChrW(234) & "n m" & ChrW(244) & "n " & ChrW(273) & ChrW(7875) & " nh" & ChrW(7853) & "n " & ChrW(273) & ChrW(432) & ChrW(7907) & "c th" & ChrW(244) & "ng tin t" & ChrW(7889) & "t nh" & ChrW(7845) & "t v" & ChrW(7873) & " m" & ChrW(225) & "y t" & ChrW(237) & "nh c" & ChrW(7911) & "a b" & ChrW(7841) & "n." & vbCrLf & "B" & ChrW(7841) & "n c" & ChrW(243) & " mu" & ChrW(7889) & "n l" & ChrW(432) & "u Log File n" & ChrW(224) & "y l" & ChrW(7841) & "i ngay b" & ChrW(226) & "y gi" _
& ChrW(7901) & " kh" & ChrW(244) & "ng?", vbYesNo, "Thông Báo", Me.hwnd) = vbYes Then cmdSaveLog_Click


End Sub

Private Sub cmdKiemtra_MouseEnter()
lblHelp.Caption = "Kie63m Tra He65 Tho61ng: Chu7o7ng tri2nh se4 ta5o 1 File Log. File Log na2y ghi la5i ca1c tho6ng so61 va2 ti2nh tra5ng ma1y ti1nh cu3a ba5n. File Log co1 kha3 na8ng d9a1nh gia1 70% ti2nh tra5ng ma1y ti1nh."
End Sub

Private Sub cmdMoRong_Click()
If sMoRong = True Then
'Thu nho lai
SaveString HKEY_CURRENT_USER, "Software\1Click", "Help", 0
Me.Height = 7320
fraHelp.Visible = False
sMoRong = False
cmdMoRong.Caption = "Mo73 Ro65ng"
Else
'Mo rong ra
SaveString HKEY_CURRENT_USER, "Software\1Click", "Help", 1
Me.Height = 8505
fraHelp.Visible = True
sMoRong = True
cmdMoRong.Caption = "Thu Nho3"
End If
End Sub

Private Sub cmdMoRong_MouseEnter()
If sMoRong = True Then lblHelp.Caption = "Ta81t Tho6ng Ba1o Hu7o71ng Da64n Na2y."
End Sub

Private Sub cmdOK_Click()
Dim IValues
Dim strKeyPath
Dim MultValueName
Dim strComputer
If List1.ListCount = 0 Then Exit Sub

strKeyPath = "SYSTEM\CurrentControlSet\Control\Session Manager"
MultValueName = "PendingFileRenameOperations"
strComputer = "."
IValues = "1"
On Error Resume Next
IValues = Array("\??\" & List1.List(0), vbNullString, "\??\" & List1.List(1), vbNullString, "\??\" & List1.List(2), vbNullString, "\??\" & List1.List(3), vbNullString, "\??\" & List1.List(4), vbNullString, "\??\" & List1.List(5), vbNullString, "\??\" & List1.List(6), vbNullString, "\??\" & List1.List(7), vbNullString, "\??\" & List1.List(8), vbNullString, "\??\" & List1.List(9), vbNullString)
Dim oReg
Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
oReg.CreateKey HKEY_LOCAL_MACHINE, strKeyPath
oReg.SetMultiStringValue HKEY_LOCAL_MACHINE, strKeyPath, MultValueName, IValues

List1.Clear

UniMsgBox ChrW$(&H58) & ChrW$(&H6F) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H2C) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1B0) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&HE1) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H66) & ChrW$(&H69) & ChrW$(&H6C) & ChrW$(&H65) & ChrW$(&H20) & ChrW$(&H76) & ChrW$(&HE0) & ChrW$(&H6F) _
& ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H61) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H73) & ChrW$(&HE1) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H78) & ChrW$(&HF3) & ChrW$(&H61) & ChrW$(&H2E) & ChrW$(&H20) & ChrW$(&H43) & ChrW$(&HE1) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H66) & ChrW$(&H69) & ChrW$(&H6C) & ChrW$(&H65) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&HE0) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H73) & ChrW$(&H1EBD) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1B0) & ChrW$(&H1EE3) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H78) & ChrW$(&HF3) _
 & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H73) & ChrW$(&H61) & ChrW$(&H75) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&H68) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&H68) & ChrW$(&H1EDF) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1ED9) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H6C) & ChrW$(&H1EA1) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H6D) & ChrW$(&HE1) & ChrW$(&H79) & ChrW$(&H2E) & vbCrLf & vbCrLf & ChrW$(&H48) & ChrW$(&HE3) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&H68) & ChrW$(&H1EDF) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1ED9) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H6C) & ChrW$(&H1EA1) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H6D) & ChrW$(&HE1) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1EC3) & ChrW$(&H20) & ChrW$(&H68) & ChrW$(&H6F) & ChrW$(&HE0) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H1EA5) & ChrW$(&H74) & ChrW$(&H2E) _
 & "", vbOKOnly, "Thông báo", Me.hwnd
End Sub

Private Sub cmdOK_MouseEnter()
lblHelp.Caption = "Xo1a File..."
End Sub

Private Sub cmdOpenFolderCha_Click()
If Not ListFind1.ListIndex = -1 Then
Shell "explorer " & GetFolderPath(ListFind1.List(ListFind1.ListIndex)), vbNormalFocus
End If

End Sub

Private Sub cmdOpenFolderCha_MouseEnter()
lblHelp.Caption = "D9i d9e61n thu7 mu5c chu71a File d9a4 cho5n."
End Sub

Private Sub cmdSaveLog_Click()
Dim sFileToSave
sFileToSave = DiaLog1.ShowSave("Text File (*.txt)|*.txt|", , "C:\", "Save Log File")
If sFileToSave <> "" Then
If Right(sFileToSave, 3) <> ".txt" Then sFileToSave = sFileToSave & ".txt"
On Error Resume Next

Open sFileToSave For Binary As #1
    Put #1, , Trim$(StrConv(txtCheck.Text, vbUnicode))
Close #1
End If
End Sub



Private Sub cmdSaveLog_MouseEnter()
lblHelp.Caption = "Lu7u la5i File Log tha2nh mo65t File va8n ba3n da5ng Notepad."
End Sub

Private Sub cmdSearch_Click()
On Error Resume Next
If Left(txtKeyWord.Text, 1) <> "*" And Right(txtKeyWord.Text, 1) <> "*" Then txtKeyWord.Text = "*" & txtKeyWord.Text & "*"
If txtPath2Find.Text <> "" Then
tmrFinding.Enabled = True
sDangTim = 1
lblStatus.Caption = "D9ang Ti2m File ..."
txtKeyWord.Enabled = False
txtPath2Find.Enabled = False
cmdSearch.Enabled = False
cmdSearchCaoCap.Enabled = False
ListFind1.Clear
sStopSearch = False
If optSearchNormal.Value = True Then
    SearchFile txtPath2Find.Text, txtKeyWord.Text, SearchBinhThuong
ElseIf optCungTenFolder.Value = True Then
     If optEXEDuoi.Value = True Then SearchFile txtPath2Find.Text, "*.exe", FileCungTenVoiThuMucChuaNo
     If optAllDuoi.Value = True Then SearchFile txtPath2Find.Text, "*.*", FileCungTenVoiThuMucChuaNo
ElseIf optSearchByIcon.Value = True Then
    UniMsgBox "Hi" & ChrW(7879) & "n t" & ChrW(7841) & "i, ch" & ChrW(7913) & "c n" & ChrW(259) & "ng nh" & ChrW(7853) & "n d" & ChrW(7841) & "ng Icon c" & ChrW(7911) & "a 1Click v" & ChrW(7851) & "n ch" & ChrW(432) & "a " & ChrW(273) & ChrW(432) & ChrW(7907) & "c ho" & ChrW(224) & "n thi" & ChrW(7879) & "n, do " & ChrW(273) & ChrW(243) & ", th" & ChrW(7901) & "i gian t" & ChrW(236) & "m ki" & ChrW(7871) & "m s" & ChrW(7869) & " l" & ChrW(226) & "u h" & ChrW(417) & "n m" & ChrW(7897) & "t t" & ChrW(253) & " (Kho" & ChrW(7843) & "ng 1 - 2 ph" & ChrW(250) & "t)." & vbCrLf _
    & ChrW(272) & ChrW(7875) & " vi" & ChrW(7879) & "c t" & ChrW(236) & "m ki" & ChrW(7871) & "m di" & ChrW(7877) & "n ra nhanh h" & ChrW(417) & "n, b" & ChrW(7841) & "n c" & ChrW(243) & " th" & ChrW(7875) & " ch" & ChrW(7885) & "n " & ChrW(273) & ChrW(7883) & "a ch" & ChrW(7881) & " t" & ChrW(236) & "m ki" & ChrW(7871) & "m cho ph" & ChrW(249) & " h" & ChrW(7907) & "p. Nh" & ChrW(7845) & "n D" & ChrW(7915) & "ng " & ChrW(273) & ChrW(7875) & " k" & ChrW(7871) & "t th" & ChrW(250) & "c vi" & ChrW(7879) & "c t" & ChrW(236) & "m ki" & ChrW(7871) & "m." & vbCrLf & "Trong khi ch" & ChrW(7901) & " 1Click t" & ChrW(236) & "m ki" & ChrW(7871) & "m, b" & ChrW(7841) & "n v" & ChrW(7851) & "n c" & ChrW(243) & " th" & ChrW(7875) & " s" & ChrW(7917) & " d" & ChrW(7909) & "ng c" & ChrW(225) & "c ch" & ChrW(7913) & "c n" & ChrW(259) & "ng kh" & ChrW(225) & "c c" & ChrW(7911) & "a 1Click b" & ChrW(236) & "nh th" & ChrW(432) & ChrW(7901) & "ng.", vbOKOnly, "!", Me.hwnd
     
     If optEXEDuoi.Value = True Then SearchFile txtPath2Find.Text, "*.exe", SearchByIcon
     If optAllDuoi.Value = True Then SearchFile txtPath2Find.Text, "*.*", SearchByIcon
        
End If
sStopSearch = True
UniMsgBox "T" & ChrW(236) & "m th" & ChrW(7845) & "y " & ListFind1.ListCount & " t" & ChrW(7853) & "p tin." & vbCrLf & vbCrLf & "Click " & ChrW(273) & ChrW(244) & "i v" & ChrW(224) & "o " & ChrW(273) & ChrW(432) & ChrW(7901) & "ng d" & ChrW(7851) & "n c" & ChrW(7911) & "a File trong danh s" & ChrW(225) & "ch " & ChrW(273) & ChrW(7875) & " " & ChrW(273) & ChrW(7871) & "n th" & ChrW(432) & " m" & ChrW(7909) & "c ch" & ChrW(7913) & "a File " & ChrW(273) & ChrW(243) & ".", vbOKOnly, "Thông báo", Me.hwnd
    txtKeyWord.Enabled = True
    txtPath2Find.Enabled = True
    cmdSearch.Enabled = True
    cmdSearchCaoCap.Enabled = True
lblStatus.Caption = "Ti2m Tha61y " & ListFind1.ListCount & " File"
tmrFinding.Enabled = False
End If


End Sub



Private Sub cmdSearch_MouseEnter()
lblHelp.Caption = "Ti2m kie61m File: Chu71c na8ng ti2m kie61m File cu3a 1Click (Nhanh ho7n 8 la62n so vo72i chu71c na8ng ti2m kie61m cu3a Windows XP)"

End Sub

Private Sub cmdSearchCaoCap_Click()
If cmdSearchCaoCap.Value = True Then
    fmNangCao.Visible = True
    txtKeyWord.Enabled = False
    txtPath2Find.Enabled = False
    cmdSearch.Enabled = False
Else
    fmNangCao.Visible = False
    txtKeyWord.Enabled = True
    txtPath2Find.Enabled = True
    cmdSearch.Enabled = True
End If
End Sub

Private Sub cmdSearchCaoCap_MouseEnter()
lblHelp.Caption = "Na6ng Cao: Thie62t la65p mo65t so61 chu71c na8ng ti2m kie61m chi tie61t va2 cu5 the63 ca1c File ho7n."
End Sub

Private Sub cmdSet_Click()



If chkDOC.Value = True Then
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions", 1
Else
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions", 0
End If


If chkPRO.Value = True Then
SaveDWORD HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer", "NoStartMenuMorePrograms", 1
Else
SaveDWORD HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer", "NoStartMenuMorePrograms", 0
End If



If chkAUTORUN.Value = True Then
SaveDWORD HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer", "NoDriveTypeAutoRun", 44
Else
SaveDWORD HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer", "NoDriveTypeAutoRun", 0
End If



If chkCMD.Value = True Then
SaveDWORD HKEY_CURRENT_USER, "Software\Policies\Microsoft\Windows\System", "DisableCMD", 1
Else
DeleteValue HKEY_CURRENT_USER, "Software\Policies\Microsoft\Windows\System", "DisableCMD"
End If



If chkCOMPUTER.Value = True Then
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoPropertiesMyComputer", 1
Else
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoPropertiesMyComputer"
End If



If chkCP.Value = True Then
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoControlPanel", 1
Else
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoControlPanel"
End If



If chkCPA.Value = True Then
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "DisallowCpl", 1
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "RestrictCpl", 1

Else
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "DisallowCpl"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "RestrictCpl"

End If



If chkCPITEM.Value = True Then
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispCpl", 1
Else
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispCpl"
End If



If chkDESKTOP.Value = True Then
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDesktop", 1
Else
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDesktop"
End If



If chkEXE.Value = True Then
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "HideFileExt", 1
Else
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "HideFileExt", 0
End If



If chkFILEMENU.Value = True Then
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "FileMenu", 1
Else
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "FileMenu"
End If


If chkHelp.Value = True Then
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSMHelp", 1
Else
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSMHelp"
End If



If chkHIDDEN.Value = True Then
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Hidden", 0
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden", 0
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "SuperHidden", 0

Else
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Hidden", 1
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden", 1
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "SuperHidden", 1
End If



If chkIEHOME.Value = True Then
SaveDWORD HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "HomePage", 1
Else
DeleteValue HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "HomePage"
End If




If chkLOGOFF.Value = True Then
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogOff", 1
Else
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogOff"
End If



If chkREG.Value = True Then
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", 1
Else
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools"
End If



If chkRIGHT.Value = True Then
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoViewContextMenu", 1
Else
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoViewContextMenu"
End If



If chkRUN.Value = True Then
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRun", 1
Else
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRun"
End If



If chkSearch.Value = True Then
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFind", 1
Else
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFind"
End If



If chkTask.Value = True Then
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr", 1
Else
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr"
End If


If chkTaskbar.Value = True Then
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSetTaskbar", 1
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSetFolders", 1

Else
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSetTaskbar"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSetFolders"

End If



If chkTRAY.Value = True Then
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoTrayContextMenu", 1
Else
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoTrayContextMenu"
End If



If chkTRAYCLOCK.Value = True Then
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "HideClock", 1
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoTrayItemsDisplay", 1

Else
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "HideClock"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoTrayItemsDisplay"

End If



If chkTURNOFF.Value = True Then
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoClose", 1
Else
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoClose"
End If



If chkUSB.Value = True Then
SaveDWORD HKEY_LOCAL_MACHINE, "SYSTEM\ControlSet001\Services\USBSTOR", "Start", 4
Else
SaveDWORD HKEY_LOCAL_MACHINE, "SYSTEM\ControlSet001\Services\USBSTOR", "Start", 3
End If



If chkWrite.Value = True Then
SaveDWORD HKEY_LOCAL_MACHINE, "SYSTEM\ControlSet001\Control\StorageDevicePolicies", "WriteProtect", 1
Else
SaveDWORD HKEY_LOCAL_MACHINE, "SYSTEM\ControlSet001\Control\StorageDevicePolicies", "WriteProtect", 0
End If

If chkWin.Value = True Then
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoWinKeys", 1
Else
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoWinKeys", 0
End If


End Sub


Private Sub cmdSet_MouseEnter()
lblHelp.Caption = "Thu75c Hie65n Thao Ta1c."
End Sub


Private Sub cmdStopSearch_Click()
sStopSearch = True

'UniMsgBox "T" & ChrW(236) & "m th" & ChrW(7845) & "y " & ListFind1.ListCount & " t" & ChrW(7853) & "p tin." & vbCrLf & vbCrLf & "Click " & ChrW(273) & ChrW(244) & "i v" & ChrW(224) & "o " & ChrW(273) & ChrW(432) & ChrW(7901) & "ng d" & ChrW(7851) & "n c" & ChrW(7911) & "a File trong danh s" & ChrW(225) & "ch " & ChrW(273) & ChrW(7875) & " " & ChrW(273) & ChrW(7871) & "n th" & ChrW(432) & " m" & ChrW(7909) & "c ch" & ChrW(7913) & "a File " & ChrW(273) & ChrW(243) & ".", vbOKOnly, "Thông báo", Me.hwnd
'    txtKeyWord.Enabled = True
'    txtPath2Find.Enabled = True
'    cmdSearch.Enabled = True
'lblStatus.Caption = "Ti2m Tha61y " & ListFind1.ListCount & " File"
'tmrFinding.Enabled = False
End Sub

Private Sub cmdTacGia_Click()
frmAbout.Show , Me
End Sub

Private Sub cmdTacGia_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdTacGia.BorderStyle = 1
End Sub

Private Sub cmdTacGia_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdTacGia.BorderStyle = 0
End Sub

Private Sub cmdTimKiemFile_Click()
FVistaUniTabStrip1.ActiveTab = 5
End Sub

Private Sub cmdTimKiemFile_MouseEnter()
lblHelp.Caption = "Ti2m kie61m File: Chu71c na8ng ti2m kie61m File cu3a 1Click (Nhanh ho7n 8 la62n so vo72i chu71c na8ng ti2m kie61m cu3a Windows XP)"
End Sub

Private Sub cmdUpdate_Click()
frmUpdate.Show , Me
End Sub

Private Sub cmdUpdate_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdUpdate.BorderStyle = 1
End Sub

Private Sub cmdUpdate_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdUpdate.BorderStyle = 0
End Sub

Private Sub ComHelp_MouseEnter()
lblHelp.Caption = "Chu71c Na8ng Kho1a Command Prompt: Mo65t so61 Virus thu7o72ng kho1a chu71c na8ng na2y d9e63 ba3o ve65 chi1nh no1, khi chu71c na8ng na2y bi5 kho1a ba5n se4 kho6ng the63 ba65t d9u7o75c Command Prompt. D9a1nh da61u d9e63 kho1a chu71c na8ng na2y, bo3 d9a1nh da61u d9e63 mo73 chu71c na8ng na2y. Nha61n nu1t thu75c hie65n d9e63 thay d9o63 co1 hie65u lu75c."

End Sub


















Private Sub ComputerHelp_MouseEnter()
lblHelp.Caption = "A63n Menu Properties Cu3a My Computer: Khi chu71c na8ng na2y bi5 kho1a, Menu Properties trong My Computer se4 bie61n ma61t. D9a1nh da61u d9e63 kho1a chu71c na8ng na2y, bo3 d9a1nh da61u d9e63 mo73 chu71c na8ng na2y. Nha61n nu1t thu75c hie65n d9e63 thay d9o63 co1 hie65u lu75c."

End Sub

Private Sub CPHelp_MouseEnter()
lblHelp.Caption = "Chu71c Na8ng Kho1a Control Panel: Mo65t so61 Virus thu7o72ng kho1a chu71c na8ng na2y d9e63 ba3o ve65 chi1nh no1, khi chu71c na8ng na2y bi5 kho1a ba5n se4 kho6ng the63 ba65t d9u7o75c Control Panel. D9a1nh da61u d9e63 kho1a chu71c na8ng na2y, bo3 d9a1nh da61u d9e63 mo73 chu71c na8ng na2y. Nha61n nu1t thu75c hie65n d9e63 thay d9o63 co1 hie65u lu75c."

End Sub


Private Sub CPITEMHelp_MouseEnter()
lblHelp.Caption = "A63n Item Trong Control Panel: Khi chu71c na8ng na2y bi5 kho1a, ca1c Icon trong Control Panel se4 bie61n ma61t. D9a1nh da61u d9e63 kho1a chu71c na8ng na2y, bo3 d9a1nh da61u d9e63 mo73 chu71c na8ng na2y. Nha61n nu1t thu75c hie65n d9e63 thay d9o63 co1 hie65u lu75c."

End Sub

Private Sub DeskHelp_MouseEnter()
lblHelp.Caption = "Chu71c Na8ng Kho1a Icon tre6n Desktop: Khi chu71c na8ng na2y bi5 kho1a ta61t ca3 ca1c Icon tre6n Desktop se4 bie61n ma61t. D9a1nh da61u d9e63 kho1a chu71c na8ng na2y, bo3 d9a1nh da61u d9e63 mo73 chu71c na8ng na2y. Nha61n nu1t thu75c hie65n d9e63 thay d9o63 co1 hie65u lu75c."

End Sub

Private Sub DocHelp_MouseEnter()
lblHelp.Caption = "A63n Menu Folder Option Tre6n Menu Explorer: La2m ma61t Menu Folder Options. D9a1nh da61u d9e63 a63n chu71c na8ng na2y, bo3 d9a1nh da61u d9e63 hie65n chu71c na8ng na2y. Nha61n nu1t thu75c hie65n d9e63 thay d9o63 co1 hie65u lu75c."

End Sub

Private Sub EXEHelp_MouseEnter()
lblHelp.Caption = "A63n D9uo6i File: Kho6ng hie63n thi5 d9uo6i File (VD: *.exe, *.txt, ...). D9a1nh da61u d9e63 a63n, bo3 d9a1nh da61u d9e63 hie65n. Nha61n nu1t thu75c hie65n d9e63 thay d9o63 co1 hie65u lu75c."

End Sub

Private Sub FileDir_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblPath.Caption = FolderDir.List(FolderDir.ListIndex) & "\" & FileDir.List(FileDir.ListIndex)
If Button = vbRightButton Then
    PopupMenu frmMenu.ttf
End If
End Sub

Private Sub FolderDir_Click()
FileDir.Path = FolderDir.List(FolderDir.ListIndex) & "\"
lblPath.Caption = FolderDir.List(FolderDir.ListIndex) & "\"
End Sub

Private Sub FolderDir_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblPath.Caption = FolderDir.List(FolderDir.ListIndex) & "\"

If Button = vbRightButton Then
    PopupMenu frmMenu.ttfo
End If
End Sub
Public Function GetFileNameNoExe(ByVal sPath As String) As String
On Error Resume Next
If InStrRev(sPath, ".") <> 0 Then
GetFileNameNoExe = Mid(sPath, InStrRev(sPath, "\") + 1, InStrRev(sPath, ".") - InStrRev(sPath, "\") - 1)
Else
GetFileNameNoExe = Mid(sPath, InStrRev(sPath, "\") + 1)
End If
End Function
Public Function GetFolderCha(ByVal sPath As String) As String
On Error Resume Next
GetFolderCha = Mid(sPath, (InStrRev(sPath, "\", InStrRev(sPath, "\") - 1)) + 1, ((InStrRev(sPath, "\") - 1) - InStrRev(sPath, "\", InStrRev(sPath, "\") - 1)))
End Function
Public Function CungTenVoiThuMucCha(ByVal sPath As String) As Boolean
On Error Resume Next
If GetFileNameNoExe(sPath) = GetFolderCha(sPath) Then CungTenVoiThuMucCha = True Else CungTenVoiThuMucCha = False

End Function



Private Sub Form_Load()
DoEvents

sStopSearch = True
FVistaUniButton13.Caption = "Hie65n Ta61t Ca3 Ca1c O63 D9i4a" & vbCrLf & "(Ke63 Ca3 O63 D9i4a Bi5 A63n)"
sMoRong = True
If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr") = 1 Then chkTask.Value = True Else chkTask.Value = False
If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools") = 1 Then chkREG.Value = True Else chkREG.Value = False
If GetDWORD(HKEY_CURRENT_USER, "Software\Policies\Microsoft\Windows\System", "DisableCMD") = 1 Then chkCMD.Value = True Else chkCMD.Value = False
If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoControlPanel") = 1 Then chkCP.Value = True Else chkCP.Value = False
If GetDWORD(HKEY_CURRENT_USER, "Software\Policies\Microsoft\Internet Explorer\Control Panel", "HomePage") = 1 Then chkIEHOME.Value = True Else chkIEHOME.Value = False
If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoTrayContextMenu") = 1 Then chkTRAY.Value = True Else chkTRAY.Value = False
If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoViewContextMenu") = 1 Then chkRIGHT.Value = True Else chkRIGHT.Value = False
If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "FileMenu") = 1 Then chkFILEMENU.Value = True Else chkFILEMENU.Value = False
If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDesktop") = 1 Then chkDESKTOP.Value = True Else chkDESKTOP.Value = False
If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSetFolders") = 1 Then chkTaskbar.Value = True Else chkTaskbar.Value = False
If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSetTaskbar") = 1 Then chkTaskbar.Value = True Else chkTaskbar.Value = False
If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoClose") = 1 Then chkTURNOFF.Value = True Else chkTURNOFF.Value = False
If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogOff") = 1 Then chkLOGOFF.Value = True Else chkLOGOFF.Value = False
If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRun") = 1 Then chkRUN.Value = True Else chkRUN.Value = False
If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFind") = 1 Then chkSearch.Value = True Else chkSearch.Value = False
If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSMHelp") = 1 Then chkHelp.Value = True Else chkHelp.Value = False
If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispCpl") = 1 Then chkCPITEM.Value = True Else chkCPITEM.Value = False
If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoTrayItemsDisplay") = 1 Then chkTRAYCLOCK.Value = True Else chkTRAYCLOCK.Value = False
If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "HideClock") = 1 Then chkTRAYCLOCK.Value = True Else chkTRAYCLOCK.Value = False
If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoPropertiesMyComputer") = 1 Then chkCOMPUTER.Value = True Else chkCOMPUTER.Value = False
If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "RestrictCpl") = 1 Then chkCPA.Value = True Else chkCPA.Value = False
If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "DisallowCpl") = 1 Then chkCPA.Value = True Else chkCPA.Value = False
If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "HideFileExt") = 1 Then chkEXE.Value = True Else chkEXE.Value = False
If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "ShowSuperHidden") = 0 Then chkHIDDEN.Value = True Else chkHIDDEN.Value = False
If GetDWORD(HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Explorer", "NoDriveTypeAutoRun") = 44 Then chkHIDDEN.Value = True Else chkHIDDEN.Value = False
If GetDWORD(HKEY_LOCAL_MACHINE, "SYSTEM\ControlSet001\Services\USBSTOR", "Start") = 4 Then chkUSB.Value = True Else chkUSB.Value = False
If GetDWORD(HKEY_LOCAL_MACHINE, "SYSTEM\ControlSet001\Control\StorageDevicePolicies", "WriteProtect") = 1 Then chkWrite.Value = True Else chkWrite.Value = False
If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoWinKeys") = 1 Then chkWin.Value = True Else chkWin.Value = False
If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoStartMenuMorePrograms") = 1 Then chkPRO.Value = True Else chkPRO.Value = False
If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions") = 1 Then chkDOC.Value = True Else chkDOC.Value = False

slVirus.Caption = ListVR.ListCount - 1

UpdateList

ListVR.AddItem "Mixa_I.exe"
ListVR.AddItem "Phimhot.exe"
ListVR.AddItem "Images.exe"
ListVR.AddItem "zPharaoh.exe"
ListVR.AddItem "Kvosoft.exe"
ListVR.AddItem "Forever.exe"
ListVR.AddItem "Shell.exe"
ListVR.AddItem "Algs.exe"
ListVR.AddItem "Amg.exe"
ListVR.AddItem "IEXPLORER.EXE"
ListVR.AddItem "pcpc.exe"
ListVR.AddItem "scvhosti.exe"
ListVR.AddItem "taskmsg.exe"
ListVR.AddItem "EV-SHUTTLE.exe"
ListVR.AddItem "Taquito.exe"
ListVR.AddItem "TiepTuc.exe"
ListVR.AddItem "Megabyte.exe"
ListVR.AddItem "sxs.exe"
ListVR.AddItem "SCVVHSOT.exe"
ListVR.AddItem "win1ogon.exe"
ListVR.AddItem "2fiy.bat"

'C:\WINDOWS\system32\drivers\klif.sys2fiy.bat


fmCaiDatChung.Left = 240
fmCaiDatChung.Top = 1680



If FileExists("C:\WINDOWS\system32\Oemlogo.bmp") = True Then
PicOemLogo.Picture = LoadPicture("C:\WINDOWS\system32\Oemlogo.bmp")
Else
PicOemLogo.AutoRedraw = True
PicOemLogo.CurrentX = (PicOemLogo.Width - Len("No Pictures Here") * 60) \ 2
PicOemLogo.CurrentY = (PicOemLogo.Height - 285) \ 2
PicOemLogo.Print "No Pictures Here"
End If


slVirus.Caption = "(" & ListVR.ListCount & ")"

'KvoSoft
GetIconFromFile "C:\WINDOWS", PicIcon1

On Error Resume Next

Dim sVers As String
sVers = GetUrlSource("http://www32.websamba.com/quangtrungsoft/version.txt")

Dim sVer As String
Dim sSion As String
sVer = Left(sVers, 1)
sSion = Right(sVers, Len(sVers) - 1)


If sVer = "3" Or sVer = "4" Or sVer = "5" Or sVer = "6" Then
    If UniMsgBox(ChrW$(&H110) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&HF3) & ChrW$(&H20) & ChrW$(&H70) & ChrW$(&H68) & ChrW$(&H69) & ChrW$(&HEA) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA3) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H31) & ChrW$(&H43) & ChrW$(&H6C) & ChrW$(&H69) & ChrW$(&H63) & ChrW$(&H6B) & ChrW$(&H20) & ChrW$(&H6D) & ChrW$(&H1EDB) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H68) & ChrW$(&H1A1) & ChrW$(&H6E) & ChrW$(&H2E) & ChrW$(&H20) & ChrW$(&H42) & ChrW$(&H1EA1) & ChrW$(&H6E) _
& ChrW$(&H20) & ChrW$(&H63) & ChrW$(&HF3) & ChrW$(&H20) & ChrW$(&H6D) & ChrW$(&H75) & ChrW$(&H1ED1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H1EA3) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H76) & ChrW$(&H1EC1) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H61) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&HE2) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H67) & ChrW$(&H69) & ChrW$(&H1EDD) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&H68) & ChrW$(&HF4) & ChrW$(&H6E) _
& ChrW$(&H67) & ChrW$(&H3F), vbYesNo, "Thông Báo", Me.hwnd) = vbYes Then Shell "explorer.exe " & sSion

End If


If GetString(HKEY_CURRENT_USER, "Software\1Click", "Help") = 1 Then
Me.Height = 8505
fraHelp.Visible = True
sMoRong = True
cmdMoRong.Caption = "Thu Nho3"
Else
Me.Height = 7320
fraHelp.Visible = False
sMoRong = False
cmdMoRong.Caption = "Mo73 Ro65ng"
End If




cmdCaiDatChung.Left = (lblSetting.Left + lblSetting.Left + lblSetting.Width) / 2 - cmdCaiDatChung.Width / 2
cmdCaiDatChung.Top = lblSetting.Top - cmdCaiDatChung.Height - 120

cmdHelp.Left = (lbllhelp.Left + lbllhelp.Left + lbllhelp.Width) / 2 - cmdHelp.Width / 2
cmdHelp.Top = lbllhelp.Top - cmdHelp.Height - 120



cmdUpdate.Left = (lblUpdate.Left + lblUpdate.Left + lblUpdate.Width) / 2 - cmdUpdate.Width / 2
cmdUpdate.Top = lblUpdate.Top - cmdUpdate.Height - 120


cmdTacGia.Left = (lblTacGia.Left + lblTacGia.Left + lblTacGia.Width) / 2 - cmdTacGia.Width / 2
cmdTacGia.Top = lblTacGia.Top - cmdTacGia.Height - 120




FileDir.Clear




GetComputerInfo

DoEvents
End Sub






Private Sub Form_Unload(Cancel As Integer)
If sStopSearch = False Then
Cancel = 1
UniMsgBox "Ch" & ChrW(432) & ChrW(417) & "ng tr" & ChrW(236) & "nh " & ChrW(273) & "ang th" & ChrW(7921) & "c hi" & ChrW(7879) & "n t" & ChrW(236) & "m ki" & ChrW(7871) & "m File, Vui l" & ChrW(242) & "ng D" & ChrW(7915) & "ng t" & ChrW(236) & "m ki" & ChrW(7871) & "m sau " & ChrW(273) & ChrW(243) & " h" & ChrW(227) & "y tho" & ChrW(225) & "t.", vbOKOnly, "!", Me.hwnd
    
Else
Dim Form As Form
   For Each Form In Forms
   Unload Form
   Set Form = Nothing
   Next Form
End If
End Sub

Private Sub FVistaUniButton1_Click()
frmWait.Show
    Dim fso As New FileSystemObject
    Dim drv As Drive
    Dim drvs As Drives
    On Error Resume Next    'in case not found, and on cd
    Set drvs = fso.Drives
    For Each drv In drvs
        DoEvents
        SetAttr drv.DriveLetter & ":\autorun.inf", vbNormal
        DeleteFile drv.DriveLetter & ":\autorun.inf"
        'Kill drv.DriveLetter & ":\autorun.inf"
    Next
    Set fso = Nothing
    Set drv = Nothing
    Set drvs = Nothing
    DoEvents

Unload frmWait
UniMsgBox _
ChrW$(&H58) & ChrW$(&H6F) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H21) & ChrW$(&H20) & ChrW$(&H54) & ChrW$(&H6F) & ChrW$(&HE0) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H42) & ChrW$(&H1ED9) & ChrW$(&H20) & ChrW$(&H46) & ChrW$(&H69) & ChrW$(&H6C) & ChrW$(&H65) & ChrW$(&H20) & ChrW$(&H41) & ChrW$(&H75) & ChrW$(&H74) & ChrW$(&H6F) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H6E) & ChrW$(&H2E) & ChrW$(&H69) & ChrW$(&H6E) & ChrW$(&H66) & ChrW$(&H20) & ChrW$(&H54) & ChrW$(&H72) & ChrW$(&H6F) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H43) & ChrW$(&HE1) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H54) & ChrW$(&H68) & ChrW$(&H1B0) & ChrW$(&H20) & ChrW$(&H4D) & ChrW$(&H1EE5) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H47) & ChrW$(&H1ED1) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H110) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H110) & ChrW$(&H1B0) & ChrW$(&H1EE3) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H58) & ChrW$(&HF3) & ChrW$(&H61) _
 & ChrW$(&H2E) & vbCrLf _
& ChrW$(&H4B) & ChrW$(&H68) & ChrW$(&H1EDF) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1ED9) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H6C) & ChrW$(&H1EA1) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H6D) & ChrW$(&HE1) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1EC3) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&HF3) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H68) & ChrW$(&H1EC3) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H1EAD) & ChrW$(&H70) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&HE1) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H1ED5) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H129) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H6D) & ChrW$(&H1ED9) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&HE1) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&HEC) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) _
& ChrW$(&H74) & ChrW$(&H68) & ChrW$(&H1B0) _
 & ChrW$(&H1EDD) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H2E), vbOKOnly, "Hoàn Thành!", Me.hwnd
End Sub

Private Sub FVistaUniButton1_MouseEnter()
lblHelp.Caption = "Khi Ba5n Kho6ng The63 Truy Ca65p Ca1c O63 D9i4a Hoa85c Khi Va2o O63 D9i4a Thi2 Hie65n Ra Cu7a So63 Open With Thi2 Co1 The63 Trong Ca1c O63 D9i4a Co1 Chu71a Autorun.inf, Chu71c Na8ng Na2y Xoa1 Toa2n Bo65 Ca1c Ta65p Tin Autorun.inf Trong Ca1c O63 D9ia4 Go61c, Nha82m D9e63 Tra1nh Vie65c Virus La6y Lan Va2 Kha81c Phu5c Lo64i Tre6n."
End Sub



Private Sub FVistaUniButton10_Click()
SaveString HKEY_CLASSES_ROOT, "exefile\shell\open\command", "", ChrW$(&H22) & ChrW$(&H25) & ChrW$(&H31) & ChrW$(&H22) & ChrW$(&H20) & ChrW$(&H25) & ChrW$(&H2A)
UniMsgBox "Xong", vbOKOnly, "Thông Báo", Me.hwnd
End Sub

Private Sub FVistaUniButton10_MouseEnter()
lblHelp.Caption = "Nha61n va2o nu1t na2y d9e63 kha81c phu5c lo64i kho6ng the63 cha5y File EXE hoa85c khi cha5y file EXE thi2 ra mo65t u71ng du5ng kha1c."
End Sub

Private Sub FVistaUniButton11_Click()
CleanReg
UniMsgBox "Ch" & ChrW(432) & ChrW(417) & "ng tr" & ChrW(236) & "nh " & ChrW(273) & ChrW(227) & " ch" & ChrW(7881) & "nh s" & ChrW(7917) & "a Registry v" & ChrW(7873) & " t" & ChrW(236) & "nh tr" & ChrW(7841) & "ng t" & ChrW(7889) & "t nh" & ChrW(7845) & "t, m" & ChrW(7885) & "i c" & ChrW(7845) & "m " & ChrW(273) & "o" & ChrW(225) & "n c" & ChrW(7911) & "a Windows " & ChrW(273) & ChrW(227) & " " & ChrW(273) & ChrW(432) & ChrW(7907) & "c m" & ChrW(7903) & ".", vbOKOnly, "Thông Báo", Me.hwnd
End Sub

Private Sub FVistaUniButton11_MouseEnter()
lblHelp.Caption = "Chu7o7ng tri2nh se4 thu74c hie65n chi3nh su73a Registry ve62 ti2nh tra5ng to61t nha61t, mo73 kho1a mo5i ca61m d9oa1n cu3a Windows, ta8ng to61c ma1y ti1nh."
End Sub

Private Sub FVistaUniButton12_Click()
SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "RegDone", "1"
UniMsgBox ChrW$(&H58) & ChrW$(&H6F) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H2C) & ChrW$(&H20) & ChrW$(&H57) & ChrW$(&H69) & ChrW$(&H6E) & ChrW$(&H64) & ChrW$(&H6F) & ChrW$(&H77) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H58) & ChrW$(&H50) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H72) & ChrW$(&H6F) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H6D) & ChrW$(&HE1) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&HED) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H1EE7) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&HF3) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA3) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H71) & ChrW$(&H75) & ChrW$(&H79) & ChrW$(&H1EC1) & ChrW$(&H6E) & ChrW$(&H2E), vbOKOnly, "Thông Báo", Me.hwnd
End Sub

Private Sub FVistaUniButton12_MouseEnter()
lblHelp.Caption = "D9a8ng Ky1 Ba3n Quye62n Cho Windows XP"
End Sub

Private Sub FVistaUniButton13_Click()
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDrives"
UniMsgBox "Xong", vbOKOnly, "Thông Báo", Me.hwnd
End Sub

Private Sub FVistaUniButton13_MouseEnter()
lblHelp.Caption = "Hie65n ca1c o63 d9i4a d9a4 bi5 a63n hoa85c d9a4 d9u7o75c ngu7o72i du2ng a63n d9i."
End Sub

Private Sub FVistaUniButton14_Click()
fmCaiDatChung.Visible = False
FVistaUniTabStrip1.Visible = True
End Sub



Private Sub FVistaUniButton14_MouseEnter()
lblHelp.Caption = "D9o1ng Cu73a So63 Ca2i D9a85t Chung, Mo73 Giao Die65n Chi1nh Cu73a Chu7o7ng Tri2nh."
End Sub

Private Sub FVistaUniButton15_MouseEnter()
lblHelp.Caption = "Mo65t so61 Virus thu7o72ng ta5o ra ca1c File co1 te6n gio61ng te6n thu7 mu5c chu71a no1 (Vi1 du5 File WINDOWS.exe na82m trong thu7 mu5c C:\WINDOWS) Nha82m d9a1nh lu72a ngu7o72i su73 du5ng Click va2o ca1c File d9o1, ta5o d9ie62u kie65n cho Virus la6y nhie64m. Chu71c na8ng na2y cho phe1p ti2m kie61m ca1c File co1 da5ng nhu7 va65y."
End Sub

Private Sub FVistaUniButton16_MouseEnter()
lblHelp.Caption = "Mo65t so61 Virus thu7o72ng ngu5y trang ba82ng ca1ch tu75 d9o63i bie63u tu7o75ng (Icon) cu3a no1 tha2nh ca1c Icon quen thuo65c (Nhu7 Icon cu3a thu7 mu5c, Icon ca1c file he65 tho61ng,...) nha82m d9a1nh lu72a ngu7o72i su73 du5ng click va2o, ta5o d9ie62u kie65n cho Virus la6y lan. Chu71c na8ng ti2m kie61m File theo Icon cu3a 1Click cho phe1p ba5n ti2m kie61m nhu74ng file co1 Icon gio61ng (hoa85c ga62n gio61ng) vo72i Icon d9a4 cho5n tru7o71c, giu1p ti2m va2 tie6u die65t trie65t d9e63 Virus."
End Sub

Private Sub FVistaUniButton2_Click()
frmWait.Show
SaveString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "default_page_url", "about:blank"
SaveString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\Main", "Search Page", "about:blank"

SaveString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\TypedURLs", "url1", "about:blank"
SaveString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\TypedURLs", "url2", "about:blank"
SaveString HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\TypedURLs", "url3", "about:blank"

UniMsgBox ChrW$(&H58) & ChrW$(&H6F) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H21) & ChrW$(&H20) & ChrW$(&H43) & ChrW$(&HE1) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H54) & ChrW$(&H68) & ChrW$(&H69) & ChrW$(&H1EBF) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H4C) & ChrW$(&H1EAD) & ChrW$(&H70) & ChrW$(&H20) & ChrW$(&H43) & ChrW$(&H1EE7) & ChrW$(&H61) _
& ChrW$(&H20) & ChrW$(&H49) & ChrW$(&H6E) & ChrW$(&H74) & ChrW$(&H65) & ChrW$(&H72) & ChrW$(&H6E) & ChrW$(&H65) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H45) & ChrW$(&H78) & ChrW$(&H70) & ChrW$(&H6C) & ChrW$(&H6F) & ChrW$(&H72) & ChrW$(&H65) & ChrW$(&H72) & ChrW$(&H20) & ChrW$(&H110) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H110) & ChrW$(&H1B0) & ChrW$(&H1EE3) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H110) & ChrW$(&H1B0) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H1EC1) & ChrW$(&H20) & ChrW$(&H54) & ChrW$(&H72) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H54) _
& ChrW$(&H68) & ChrW$(&HE1) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H4D) & ChrW$(&H1EB7) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H110) & ChrW$(&H1ECB) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H2E) & ChrW$(&H20) & ChrW$(&H4E) & ChrW$(&H68) & ChrW$(&H1EA5) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H4E) & ChrW$(&HFA) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H22) & ChrW$(&H110) & ChrW$(&HF3) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H22) & ChrW$(&H20) & ChrW$(&H110) & ChrW$(&H1EC3) & ChrW$(&H20) & ChrW$(&H48) & ChrW$(&H6F) & ChrW$(&HE0) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H54) & ChrW$(&H1EA5) & ChrW$(&H74) & ChrW$(&H2E), vbOKOnly, "Hoàn Thành", Me.hwnd
Unload frmWait
End Sub

Private Sub FVistaUniButton2_MouseEnter()
lblHelp.Caption = "Khi su73 du5ng chu71c na8ng na2y, ca1c thie61t la65p cu3a Internet Explorer se4 tro73 ve62 tra5ng tha1i ma8c d9inh nhu7 lu1c mo71i ca2i va2o ma1y."
End Sub

Private Function SearchFile(Path, FileName, ByVal sKieuTim As KieuSearch)

   Dim FP As FILE_PARAMS  'holds search parameters
   Dim tstart As Single   'timer var for this routine only
   Dim tend As Single     'timer var for this routine only
   With FP
      .sFileRoot = Path       'start path
      .sFileNameExt = FileName    'file type of interest
      .bRecurse = 1 ' Check1.Value = 1  '1 = recursive search
   End With
   tstart = GetTickCount()
   Call SearchForFiles(FP, sKieuTim)
   tend = GetTickCount()
End Function


Private Sub GetFileInformation(FP As FILE_PARAMS, KieuTim)
DoEvents
   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
   Dim sPath As String
   Dim sRoot As String
   Dim sTmp As String
   Dim SKetQua As String
   sRoot = QualifyPath(FP.sFileRoot)
   sPath = sRoot & FP.sFileNameExt
   hFile = FindFirstFile(sPath, WFD)
   If hFile <> INVALID_HANDLE_VALUE Then
      Do
         If Not (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = _
                 FILE_ATTRIBUTE_DIRECTORY Then
            FP.Count = FP.Count + 1
            sTmp = TrimNull(WFD.cFileName)
            SKetQua = sRoot & sTmp
            
            If KieuTim = 0 Then
            ListFind1.AddItem SKetQua
            ElseIf KieuTim = 1 Then
            'ListFind1.AddItem SKetQua
            If CungTenVoiThuMucCha(SKetQua) = True Then ListFind1.AddItem SKetQua
            ElseIf KieuTim = 2 Then
                GetIconFromFile SKetQua, frmMain.PicIcon2
                Dim sGiongNhau
                    sGiongNhau = SoSanhPic(PicIcon1, PicIcon2)
                    If sGiongNhau >= 80 And frmMain.chkIconChinhXac.Value = True Then
                        ListFind1.AddItem SKetQua
                    ElseIf sGiongNhau >= 50 And frmMain.chkIconGanGiong.Value = True Then
                        ListFind1.AddItem SKetQua
                    ElseIf sGiongNhau >= 30 And frmMain.chkIconTimHet.Value = True Then
                        ListFind1.AddItem SKetQua
                    End If
                
            'ListFind1.AddItem SKetQua & "icon"
            End If
            
          If sStopSearch = True Then Exit Sub
            '*********************************************
         End If
      Loop While FindNextFile(hFile, WFD)
      hFile = FindClose(hFile)
   End If
DoEvents
End Sub

Private Sub SearchForFiles(FP As FILE_PARAMS, tKieuTim)
  'local working variables
   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
   Dim sPath As String
   Dim sRoot As String
   Dim sTmp As String
   sRoot = QualifyPath(FP.sFileRoot)
   sPath = sRoot & "*.*"
   hFile = FindFirstFile(sPath, WFD)
   If hFile <> INVALID_HANDLE_VALUE Then
      Call GetFileInformation(FP, tKieuTim)
      Do
         If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
            If FP.bRecurse Then
               If Asc(WFD.cFileName) <> vbDot Then
                  FP.sFileRoot = sRoot & TrimNull(WFD.cFileName)
                  If sStopSearch = True Then Exit Sub
                  Call SearchForFiles(FP, tKieuTim)
               End If
            End If
         End If
      Loop While FindNextFile(hFile, WFD)
      hFile = FindClose(hFile)
   End If
End Sub



Private Function QualifyPath(sPath As String) As String
   If Right$(sPath, 1) <> "\" Then
      QualifyPath = sPath & "\"
   Else
      QualifyPath = sPath
   End If
End Function


Private Function TrimNull(startstr As String) As String
   Dim pos As Integer
   pos = InStr(startstr, Chr$(0))
   If pos Then
      TrimNull = Left$(startstr, pos - 1)
      Exit Function
   End If
   TrimNull = startstr
End Function
Private Sub FVistaUniButton3_Click()
frmWait.Show
Dim str
Dim str2
Dim fso As New FileSystemObject
    Dim drv As Drive
    Dim drvs As Drives
    On Error Resume Next
    Set drvs = fso.Drives
    For Each drv In drvs
        DoEvents
        str = "cmd /c md \\?\" & drv.DriveLetter & ":\autorun.inf"
        str2 = "cmd /c md \\?\" & drv.DriveLetter & ":\autorun.inf\.1Click.QuangTrung."
        Shell str, vbHide
        Shell str2, vbHide
        SetAttr drv.DriveLetter & ":\autorun.inf", vbHidden + vbSystem + vbReadOnly
        Shell "cmd /c attrib " & drv.DriveLetter & ":\autorun.inf +s +h", vbHide
        
        ExtracIcon drv.DriveLetter & ":\autorun.inf\1Click-Protect.ico"
        SetAttr drv.DriveLetter & ":\autorun.inf\1Click-Protect.ico", vbSystem + vbReadOnly + vbHidden
        CreateFolderIcon drv.DriveLetter & ":\autorun.inf"
        'Kill drv.DriveLetter & ":\autorun.inf"
    Next
    Set fso = Nothing
    Set drv = Nothing
    Set drvs = Nothing
    DoEvents

EndTask "cmd.exe"
Unload frmWait
UniMsgBox ChrW$(&H110) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H54) & ChrW$(&H68) & ChrW$(&H69) & ChrW$(&H1EBF) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H4C) & ChrW$(&H1EAD) & ChrW$(&H70) & ChrW$(&H20) & ChrW$(&H43) & ChrW$(&H68) & ChrW$(&H1EE9) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H4E) & ChrW$(&H103) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H43) & ChrW$(&H68) & ChrW$(&H1ED1) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H41) & ChrW$(&H75) & ChrW$(&H74) & ChrW$(&H6F) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H6E) & ChrW$(&H2E) & ChrW$(&H69) & ChrW$(&H6E) & ChrW$(&H66) & ChrW$(&H2E) _
& vbCrLf _
& ChrW$(&H4E) & ChrW$(&H1EBF) & ChrW$(&H75) & ChrW$(&H20) & ChrW$(&H42) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H4D) & ChrW$(&H75) & ChrW$(&H1ED1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H47) & ChrW$(&H1EE1) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1ECF) & ChrW$(&H20) & ChrW$(&H43) & ChrW$(&H68) & ChrW$(&H1EE9) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H4E) _
& ChrW$(&H103) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H4E) & ChrW$(&HE0) & ChrW$(&H79) & ChrW$(&H2C) & ChrW$(&H20) & ChrW$(&H48) & ChrW$(&HE3) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H43) & ChrW$(&H6C) & ChrW$(&H69) & ChrW$(&H63) & ChrW$(&H6B) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&HE0) & ChrW$(&H6F) & ChrW$(&H20) & ChrW$(&H4E) & ChrW$(&HFA) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H22) & ChrW$(&H47) & ChrW$(&H1EE1) & ChrW$(&H22) & ChrW$(&H20) & ChrW$(&H42) & ChrW$(&HEA) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H43) & ChrW$(&H1EA1) & ChrW$(&H6E) _
 & ChrW$(&H68) & ChrW$(&H2E), vbOKOnly, "Hoàn Thành", Me.hwnd
End Sub

Private Sub FVistaUniButton3_MouseEnter()
lblHelp.Caption = "Ta5o ca1c Autorun gia3 va2o ca1c o63 d9ia4 d9e63 tra1nh ca1c Virus la6y nhie64m va2o ma1y ti1nh."
End Sub

Private Sub FVistaUniButton4_Click()
frmWait.Show
On Error Resume Next
Kill "C:\WINDOWS\Prefetch\*.*"

SaveDWORD HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer", "AlwaysUnloadDll", 1
SaveString HKEY_CURRENT_USER, "Control Panel/Desktop", "MenuShowDelay", "00000000"
SaveDWORD HKEY_CURRENT_USER, "Control Panel/Desktop", "AutoEndTask", 1
SaveDWORD HKEY_CURRENT_USER, "Control Panel/Desktop", "HungAppTimeout", 200
SaveDWORD HKEY_CURRENT_USER, "Control Panel/Desktop", "WaitToKillAppTimeOut", 200
SaveDWORD HKEY_CURRENT_USER, "Control Panel/Desktop", "WaitToKillServicesOut", 200


Unload frmWait
UniMsgBox ChrW$(&H4F) & ChrW$(&H4B) & ChrW$(&H21) & ChrW$(&H20) & ChrW$(&H42) & ChrW$(&HE2) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H67) & ChrW$(&H69) & ChrW$(&H1EDD) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H1ED1) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1ED9) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H1EAF) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H76) & ChrW$(&HE0) _
& ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&H68) & ChrW$(&H1EDF) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1ED9) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H6D) & ChrW$(&HE1) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&HED) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H1EE7) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H61) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H103) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H6C) _
 & ChrW$(&HEA) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H72) & ChrW$(&H1EA5) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H69) & ChrW$(&H1EC1) & ChrW$(&H75) & ChrW$(&H21), vbOKOnly, "Hoàn Thành", Me.hwnd

End Sub


Private Sub FVistaUniButton4_MouseEnter()
lblHelp.Caption = "Khi cho5n chu71c na8ng na2y, chu7o7ng tri2nh se4 thie61t la65p ca1c tho6ng so61 trong Registry d9e63 ma1y ti1nh hoa5t d9o65ng to61t nha61t."
End Sub

Private Sub FVistaUniButton5_Click()

'C:\WINDOWS\system32\drivers\etc\hosts
'DeleteFile "C:\WINDOWS\system32\drivers\etc\hosts"
SetAttr "C:\WINDOWS\system32\drivers\etc\hosts", vbNormal
DeleteFile "C:\WINDOWS\system32\drivers\etc\hosts"
CreateTextFile "C:\WINDOWS\system32\drivers\etc\hosts", "127.0.0.1       localhost"

UniMsgBox "Xong", vbOKOnly, "Thông Báo", Me.hwnd
End Sub



Private Sub FVistaUniButton5_MouseEnter()
lblHelp.Caption = "Khi ba5n va2o mo65t trang web na2o d9o1 ma2 la5i bi5 chuye63n hu7o71ng sang mo65t trang kha1c (VD: va2o http://google.com.vn thi2 bi5 chuye63n sang trang http://xxx.tk) thi2 ra61t co1 the63 ta65p tin Hosts d9a4 bi5 thay d9o63i. Khi cho5n chu71c na8ng na2y, chu7o7ng tri2nh se4 su73a chu74a ta65p tin Hosts la5i nguye6n da5ng nhu7 ban d9a62u."
End Sub

Private Sub FVistaUniButton6_Click()
SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Userinit", "C:\WINDOWS\system32\userinit.exe,"
SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", "Explorer.exe"
UniMsgBox "Xong", vbOKOnly, "Thông Báo", Me.hwnd
End Sub

Private Sub FVistaUniButton6_MouseEnter()
lblHelp.Caption = "Su7a3 chu74a lo64i kho6ng the63 Log On va2o ma1y ti1nh."
End Sub

Private Sub FVistaUniButton7_Click()
frmWait.Show
Dim sDriver, i
sDriver = "ZXCVBNMLKJHGFDSAQWERTYUIOP"
For i = 1 To Len(sDriver)
    On Error Resume Next
Shell "cmd /c rd \\?\" & Mid(sDriver, i, 1) & ":\autorun.inf\.1Click.QuangTrung.", vbHide
Shell "cmd /c del " & Mid(sDriver, i, 1) & ":\autorun.inf", vbHide
DoEvents
EndTask "cmd.exe"
'RmDir Mid(sDriver, i, 1) & ":\autorun.inf"
Next i
Unload frmWait
UniMsgBox ChrW$(&H110) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H47) & ChrW$(&H1EE1) & ChrW$(&H20) & ChrW$(&H42) & ChrW$(&H1ECF) & ChrW$(&H20) & ChrW$(&H43) & ChrW$(&H68) & ChrW$(&H1EE9) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H4E) & ChrW$(&H103) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H43) & ChrW$(&H68) & ChrW$(&H1ED1) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H41) & ChrW$(&H75) & ChrW$(&H74) & ChrW$(&H6F) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H6E) & ChrW$(&H2E) & ChrW$(&H69) & ChrW$(&H6E) & ChrW$(&H66), vbOKOnly, "Thông Báo", Me.hwnd
End Sub


Private Sub FVistaUniButton7_MouseEnter()
lblHelp.Caption = "Go74 bo3 chu71c na8ng ta5o Autorun gia3. Sau khi nha61n nu1t na8ng na2y, ca1c thu7 mu5c Autorun.inf va64n co2n trong ca1c o63 d9ia4, nhu7ng ba5n co1 the63 xo1a no1 ba82ng tay."
End Sub

Private Sub FVistaUniButton8_Click()
Shell "taskkill /f /im explorer.exe"
tmrStart.Enabled = True
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHelp.Caption = "A"
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHelp.Caption = "B"
End Sub

Private Sub FVistaUniButton8_MouseEnter()
lblHelp.Caption = "Ba81t d9a62u nga8n File"
End Sub

Private Sub FVistaUniButton9_Click()

Dim sFileName

sFileName = DiaLog1.ShowOpen("EXE File (*.exe)", , "C:\", "Select File")
    
    
        Dim i, X, a
        a = sFileName
    For i = 1 To Len(a)
        If Mid(a, i, 1) = "\" Then
            X = i
        End If
    Next i
    txtFileName.Text = Right(a, Len(a) - X)
    
    
  
ER:
End Sub

Private Sub FVistaUniButton9_MouseEnter()
lblHelp.Caption = "Cho5n File"
End Sub









Private Sub FVistaUniTabStrip1_BeforeTabSwitch(ByVal iNewActiveTab As Integer, bCancel As Boolean)
If iNewActiveTab = 0 Then lblHelp.Caption = "Chu71c Na8ng Na2y Cho Phe1p Ba5n Mo73 Hoa85c Kho1a Mo65t So61 Chu71c Na8ng Ca62n Thie61t Cu3a Windows, Mo65t So61 Virus Thu7o72ng Kho1a Ca1c Chu71c Na8ng Na2y D9e63 Ba3o Ve65 Chi1nh No1, D9a1nh Da61u Va2o Ca1c Khung Lu75a Cho5n D9e63 Kho1a Va2 Bo3 D9a1nh Da61u D9e63 Mo73 Kho1a. Nha61n Nu1t Thu75c Hie65n D9e63 Thay D9o63i Co1 Hie65u Lu75c."
If iNewActiveTab = 4 Then FileDir.Path = "X:\"

End Sub



Private Sub HelpHelp_MouseEnter()
lblHelp.Caption = "A63n Chu71c Na8ng Help: Khi chu71c na8ng na2y bi5 kho1a, nu1t Help tre6n Menu Start se4 bie61n ma61t. D9a1nh da61u d9e63 kho1a chu71c na8ng na2y, bo3 d9a1nh da61u d9e63 mo73 chu71c na8ng na2y. Nha61n nu1t thu75c hie65n d9e63 thay d9o63 co1 hie65u lu75c."

End Sub

Private Sub HIDECPHelp_Click()
lblHelp.Caption = "A63n Control Panel"
End Sub

Private Sub HIDEHelp_MouseEnter()
lblHelp.Caption = "A63n File: Kho6ng hie63n thi5 d9uo6i File a63n va2 ca1c File he65 tho61ng. D9a1nh da61u d9e63 a63n, bo3 d9a1nh da61u d9e63 hie65n. Nha61n nu1t thu75c hie65n d9e63 thay d9o63 co1 hie65u lu75c."

End Sub

Private Sub IEHelp_MouseEnter()
lblHelp.Caption = "Chu71c Na8ng Kho1a Internet Explorer Home Pages: Khi chu71c na8ng na2y bi5 kho1a ba5n se4 kho6ng the63 thay d9o63 Trang Chu3 cu3a Internet Explorer. D9a1nh da61u d9e63 kho1a chu71c na8ng na2y, bo3 d9a1nh da61u d9e63 mo73 chu71c na8ng na2y. Nha61n nu1t thu75c hie65n d9e63 thay d9o63 co1 hie65u lu75c."

End Sub

Private Sub List1_Click()
lblFile.Caption = List1.List(List1.ListIndex)
End Sub

Private Sub ListCam_Click()
Inde = ListCam.ListIndex
lblHelp.Caption = "Ca1c File Bi5 Ca61m"
End Sub



Public Function GetFileName(ByVal sPath As String) As String
GetFileName = Mid(sPath, InStrRev(sPath, "\") + 1)
End Function
Public Function GetFolderPath(ByVal sPath As String) As String
GetFolderPath = Left(sPath, InStrRev(sPath, "\") - 1)
End Function

Private Sub ListFind1_Click()
If ListFind1.ListIndex <> -1 Then
lblInfoSize.Caption = ""
lblStatus.Caption = ListFind1.List(ListFind1.ListIndex)
GetIconFromFile lblStatus.Caption, PicInfoIcon
lblInfoSize.Caption = FileLen(lblStatus.Caption) & " Bytes"
End If
End Sub

Private Sub ListFind1_DbClick()
If Not ListFind1.ListIndex = -1 Then
Shell "explorer " & GetFolderPath(ListFind1.List(ListFind1.ListIndex)), vbNormalFocus
End If
End Sub

Private Sub ListVR_Click()
lblHelp.Caption = "Ne61u ba5n bie61t (hoa85c nghi ngo72) ra82ng ma1y ti1nh cu73a ba5n d9ang bi5 nhie64m mo65t trong ca1c loa5i Virus co1 trong danh sa1ch tre6n, khi d9o1 ban5 chi3 ca62n nha61p va2o te6n Virus ma2 ba5n nghi ngo72 la2 ba5n bi5 nhie64m, chu7o7ng tri2nh se4 to61ng co63 no1 ra kho3i ma1y ti1nh cu73a ba5n ngay la65p tu71c, d9a3m ba3o sa5ch ta65n go61c."
cmdDiet.Enabled = True

End Sub

Private Sub LogOffHelp_MouseEnter()
lblHelp.Caption = "A63n Nu1t Log Off: Khi chu71c na8ng na2y bi5 kho1a, nu1t Log Off tre6n Menu Start se4 bie61n ma61t. D9a1nh da61u d9e63 kho1a chu71c na8ng na2y, bo3 d9a1nh da61u d9e63 mo73 chu71c na8ng na2y. Nha61n nu1t thu75c hie65n d9e63 thay d9o63 co1 hie65u lu75c."

End Sub

Private Sub MENUHelp_MouseEnter()
lblHelp.Caption = "Chu71c Na8ng Kho1a Menu File: Khi chu71c na8ng na2y bi5 kho1a ba5n se4 kho6ng the63 cho5n Menu File trong ca1c cu73a so63 la2m vie65c. D9a1nh da61u d9e63 kho1a chu71c na8ng na2y, bo3 d9a1nh da61u d9e63 mo73 chu71c na8ng na2y. Nha61n nu1t thu75c hie65n d9e63 thay d9o63 co1 hie65u lu75c."
End Sub

Private Sub OffHelp_MouseEnter()
lblHelp.Caption = "A63n Nu1t Turn Off: Khi chu71c na8ng na2y bi5 kho1a, nu1t Turn Off tre6n Menu Start se4 bie61n ma61t. D9a1nh da61u d9e63 kho1a chu71c na8ng na2y, bo3 d9a1nh da61u d9e63 mo73 chu71c na8ng na2y. Nha61n nu1t thu75c hie65n d9e63 thay d9o63 co1 hie65u lu75c."

End Sub

Private Sub optCungTenFolder_Click()

    lblLoaiDuoi.Enabled = True
    optAllDuoi.Value = True
    optEXEDuoi.Enabled = True
    optAllDuoi.Enabled = True

    lblHelpClickIcon.Enabled = False
    chkIconChinhXac.Enabled = False
    chkIconGanGiong.Enabled = False
    frmMain.chkIconTimHet.Enabled = False
    frmMain.PicIcon1.Enabled = False
    chkIconChinhXac.Value = False
    chkIconGanGiong.Value = False
    chkIconTimHet.Value = False
    
End Sub

Private Sub optSearchByIcon_Click()
    lblLoaiDuoi.Enabled = True
    optAllDuoi.Value = True
    optEXEDuoi.Enabled = True
    optAllDuoi.Enabled = True
   
    
    lblHelpClickIcon.Enabled = True
    chkIconChinhXac.Enabled = True
    chkIconGanGiong.Enabled = True
    frmMain.chkIconTimHet.Enabled = True
    frmMain.PicIcon1.Enabled = True
    chkIconChinhXac.Value = True
    
    
End Sub

Private Sub optSearchNormal_Click()
    lblLoaiDuoi.Enabled = False
    optEXEDuoi.Enabled = False
    optAllDuoi.Enabled = False
    optEXEDuoi.Value = False
    optAllDuoi.Value = False
    
    
    lblHelpClickIcon.Enabled = False
    chkIconChinhXac.Enabled = False
    chkIconGanGiong.Enabled = False
    frmMain.chkIconTimHet.Enabled = False
    frmMain.PicIcon1.Enabled = False
    chkIconChinhXac.Value = False
    chkIconGanGiong.Value = False
    chkIconTimHet.Value = False
    
End Sub

Private Sub PicIcon1_Click()
Dim sPict As String
sPict = DiaLog1.ShowOpen("All File", , "C:\", "Select File To Make Icon")
If sPict <> "" Then
    GetIconFromFile sPict, PicIcon1
End If
End Sub

Private Sub ProgramHelp_MouseEnter()
lblHelp.Caption = "A63n Nu1t All Program Tre6n Menu Start: La2m ma61t nu1t All Program trong Start Menu Classic. D9a1nh da61u d9e63 kho1a chu71c na8ng na2y, bo3 d9a1nh da61u d9e63 mo73 chu71c na8ng na2y. Nha61n nu1t thu75c hie65n d9e63 thay d9o63 co1 hie65u lu75c."

End Sub

Private Sub RegHelp_MouseEnter()
lblHelp.Caption = "Chu71c Na8ng Kho1a Registry: Mo65t so61 Virus thu7o72ng kho1a chu71c na8ng na2y d9e63 ba3o ve65 chi1nh no1, khi chu71c na8ng na2y bi5 kho1a ba5n se4 kho6ng the63 ba65t d9u7o75c Registry. D9a1nh da61u d9e63 kho1a chu71c na8ng na2y, bo3 d9a1nh da61u d9e63 mo73 chu71c na8ng na2y. Nha61n nu1t thu75c hie65n d9e63 thay d9o63 co1 hie65u lu75c."

End Sub


Private Sub RIGHTHelp_MouseEnter()
lblHelp.Caption = "Chu71c Na8ng Kho1a Chuo65t Pha3i: Khi chu71c na8ng na2y bi5 kho1a ba5n se4 kho6ng the63 Click Chuo65t Pha3i. D9a1nh da61u d9e63 kho1a chu71c na8ng na2y, bo3 d9a1nh da61u d9e63 mo73 chu71c na8ng na2y. Nha61n nu1t thu75c hie65n d9e63 thay d9o63 co1 hie65u lu75c."
End Sub



Private Sub RunHelp_MouseEnter()
lblHelp.Caption = "A63n Chu71c Na8ng Run...: Khi chu71c na8ng na2y bi5 kho1a, nu1t Run... tre6n Menu Start se4 bie61n ma61t. D9a1nh da61u d9e63 kho1a chu71c na8ng na2y, bo3 d9a1nh da61u d9e63 mo73 chu71c na8ng na2y. Nha61n nu1t thu75c hie65n d9e63 thay d9o63 co1 hie65u lu75c."

End Sub

Private Sub StartHelp_MouseEnter()
lblHelp.Caption = "A63n Chu71c Na8ng Search: Khi chu71c na8ng na2y bi5 kho1a, nu1t Search tre6n Menu Start se4 bie61n ma61t. D9a1nh da61u d9e63 kho1a chu71c na8ng na2y, bo3 d9a1nh da61u d9e63 mo73 chu71c na8ng na2y. Nha61n nu1t thu75c hie65n d9e63 thay d9o63 co1 hie65u lu75c."

End Sub



Private Sub TaskbarHelp_MouseEnter()
lblHelp.Caption = "Chu71c Na8ng Kho1a Ca2i D9a85t Taskbar Va2 Menu Start: Khi chu71c na8ng na2y bi5 kho1a ta61t ca3 ca1c Icon tre6n Desktop se4 bie61n ma61t. D9a1nh da61u d9e63 kho1a chu71c na8ng na2y, bo3 d9a1nh da61u d9e63 mo73 chu71c na8ng na2y. Nha61n nu1t thu75c hie65n d9e63 thay d9o63 co1 hie65u lu75c."

End Sub

Private Sub TaskHelp_MouseEnter()
lblHelp.Caption = "Chu71c Na8ng Kho1a Task Manager: Mo65t so61 Virus thu7o72ng kho1a chu71c na8ng na2y d9e63 ba3o ve65 chi1nh no1, khi chu71c na8ng na2y bi5 kho1a ba5n se4 kho6ng the63 ba65t d9u7o75c Task Manager. D9a1nh da61u d9e63 kho1a chu71c na8ng na2y, bo3 d9a1nh da61u d9e63 mo73 chu71c na8ng na2y. Nha61n nu1t thu75c hie65n d9e63 thay d9o63 co1 hie65u lu75c."
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If GetString(HKEY_CURRENT_USER, "Software\1Click", "FirstTime") = "0" Then
SaveString HKEY_CURRENT_USER, "Software\1Click", "FirstTime", "1"
frmHelpMe.Show
End If
Timer1.Enabled = False
End Sub




Private Sub tmrFinding_Timer()
sDangTim = sDangTim + 1
If sDangTim = 2 Then
lblStatus.Caption = "D9ang Ti2m File ... |"
ElseIf sDangTim = 3 Then
lblStatus.Caption = "D9ang Ti2m File ... /"
ElseIf sDangTim = 4 Then
lblStatus.Caption = "D9ang Ti2m File ... ---"
ElseIf sDangTim = 5 Then
lblStatus.Caption = "D9ang Ti2m File ... \"
ElseIf sDangTim = 6 Then
lblStatus.Caption = "D9ang Ti2m File ... |"
sDangTim = 1
End If
End Sub

Private Sub tmrStart_Timer()
Unload frmWait
UniMsgBox ChrW$(&H110) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H54) & ChrW$(&H68) & ChrW$(&H1EF1) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H48) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H58) & ChrW$(&H6F) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H21), vbOKOnly, "Thông Báo", Me.hwnd
Shell "explorer.exe"

tmrStart.Enabled = False
End Sub
Private Sub UpdateList()

Dim i
For i = 0 To 100
If GetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun", i) <> "" Then
ListCam.AddItem GetString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\DisallowRun", i)
End If
Next i
End Sub


Private Sub TrayHelp_MouseEnter()
lblHelp.Caption = "Chu71c Kho6ng Cho Click Chuo65t Pha3i Taskbar: Mo65t so61 Virus thu7o72ng kho1a chu71c na8ng na2y d9e63 ba3o ve65 chi1nh no1, khi chu71c na8ng na2y bi5 kho1a ba5n se4 kho6ng the63 Click chuo65t pha3i va2o thanh Taskbar. D9a1nh da61u d9e63 kho1a chu71c na8ng na2y, bo3 d9a1nh da61u d9e63 mo73 chu71c na8ng na2y. Nha61n nu1t thu75c hie65n d9e63 thay d9o63 co1 hie65u lu75c."
End Sub



Private Sub txtFileName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHelp.Caption = "Te6n File"
End Sub


Private Sub txtPath2Find_Click()
Dim sPathFind
sPathFind = BrowserFolder("Select Folder To Find")
If sPathFind <> "" Then txtPath2Find.Text = sPathFind
End Sub

Private Sub USBHelp_MouseEnter()
lblHelp.Caption = "Ca61m Kho6ng Cho Su73 Du5ng USB Ta5i Ma1y Ti1nh Na2y."
End Sub

Private Sub WinHelp_MouseEnter()
lblHelp.Caption = "Chu71c Na8ng Kho1a Phi1m Windows: Khi chu71c na8ng na2y bi5 kho1a, phi1m Windows se4 bi5 ma61t ta1c du5ng. D9a1nh da61u d9e63 kho1a chu71c na8ng na2y, bo3 d9a1nh da61u d9e63 mo73 chu71c na8ng na2y. Nha61n nu1t thu75c hie65n d9e63 thay d9o63 co1 hie65u lu75c."

End Sub

Private Sub WriteHelp_MouseEnter()
lblHelp.Caption = "Kho6ng Cho Ghi Du74 Lie65u Va2o USB (Chi3 Co1 Ta1c Du5ng Ta5i Ma1y Ti1nh Na2y)."
End Sub

Function GetOS()
Dim strComputer, strWMIOS
strComputer = "."
Dim objWmiService: Set objWmiService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Dim strOsQuery: strOsQuery = "Select * from Win32_OperatingSystem"
Dim colOperatingSystems: Set colOperatingSystems = objWmiService.ExecQuery(strOsQuery)
Dim objOs
Dim strOsVer

    For Each objOs In colOperatingSystems
        strWMIOS = objOs.Caption & " " & objOs.Version
    Next
GetOS = strWMIOS
End Function
Function GetUser()
    GetUser = Environ$("username")
End Function
Function GetComputer()
    Dim dwlen As Long
    Dim strString As String
    dwlen = MAX_COMPUTERNAME_LENGTH + 1
    strString = String(dwlen, "X")
    GetComputerName strString, dwlen
    strString = Left(strString, dwlen)
    GetComputer = strString
End Function
Public Function GetDriveTotal() As String
    Dim TotalSize
    Dim fso As New FileSystemObject
    Dim drv As Drive
    Dim drvs As Drives
    TotalSize = 0
    On Error Resume Next    'in case not found, and on cd
    Set drvs = fso.Drives
    For Each drv In drvs
        Dim ABC
        Set ABC = fso.GetDrive(fso.GetDriveName(drv.DriveLetter & ":\"))
        TotalSize = TotalSize + ABC.TotalSize
    Next
    Set fso = Nothing
    Set drv = Nothing
    Set drvs = Nothing
    DoEvents

GetDriveTotal = Round(TotalSize / (10 ^ 9), 3)
End Function
Public Function GetRAMTotal() As String
   Call GlobalMemoryStatus(memInfo)
        GetRAMTotal = Round(memInfo.dwTotalPhys / 1024 / 1024, 3) & " MB"
End Function
Public Function AutorunProtect() As String
On Error Resume Next

SetAttr "C:\autorun.inf", vbNormal
If IsFolder("C:\autorun.inf") = True Then
    AutorunProtect = "D9a4 D9u7o75c Ba3o Ve65"
Else
    AutorunProtect = "Chu7a D9u7o75c Ba3o Ve65"
End If
SetAttr "C:\autorun.inf", vbHidden + vbSystem + vbReadOnly
End Function
Public Function IsFolder(PathFile As String) As Boolean
    'Khong ton tai
    If Dir(PathFile) = "" And Dir(PathFile, vbDirectory) = "" Then
        
    Else
        
        'Day la File
        If Dir(PathFile) <> "" Then
            IsFolder = False
        Else 'Day la Folder
            IsFolder = True
        End If
    End If
End Function
Public Function AutorunView() As String
On Error Resume Next
Dim sString
Dim sYes As Boolean
sString = "Co1 Autorun O73: "
SetAttr "C:\autorun.inf", vbNormal
If FileExists("C:\autorun.inf") = True Then
sString = sString & "|C:\"
sYes = True
End If
SetAttr "D:\autorun.inf", vbNormal
If FileExists("D:\autorun.inf") = True Then
sString = sString & "|D:\"
sYes = True
End If
SetAttr "E:\autorun.inf", vbNormal
If FileExists("E:\autorun.inf") = True Then
sString = sString & "|E:\"
sYes = True
End If
SetAttr "F:\autorun.inf", vbNormal
If FileExists("F:\autorun.inf") = True Then
sString = sString & "|F:\"
sYes = True
End If
If sYes = False Then sString = "Kho6ng Pha1t Hie65n Autorun."

SetAttr "C:\autorun.inf", vbHidden + vbSystem + vbReadOnly

AutorunView = sString
End Function

Public Function CheckTaskManager() As String
If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr") = 1 Then CheckTaskManager = "D9ang bi5 kho1a." Else CheckTaskManager = "D9ang d9u7o75c mo73."
End Function
Public Function CheckRegistry() As String
If GetDWORD(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools") = 1 Then CheckRegistry = "D9ang bi5 kho1a." Else CheckRegistry = "D9ang d9u7o75c mo73."
End Function

'SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Userinit", "C:\WINDOWS\system32\userinit.exe"
'SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", "explorer.exe"
Public Function CheckLogOnErr() As String
If GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Userinit") <> "C:\WINDOWS\system32\userinit.exe," Or GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell") <> "Explorer.exe" Then CheckLogOnErr = "D9ang Bi5 Lo64i." Else CheckLogOnErr = "Kho6ng Bi5 Lo64i."
End Function
Function GetMemoryInfo()

  DoEvents
  GlobalMemoryStatus memoryInfo
    Dim Totp1
    Dim Availp1
    Dim pcent
    Dim lastpcent
    Dim lastTot
  Totp1 = Int(memoryInfo.dwTotalPhys / 1044032 * 10 + 0.5) / 10
  Availp1 = Int(memoryInfo.dwAvailPhys / 1044032 * 10 + 0.5) / 10
  pcent = Int(Availp1 / Totp1 * 100)
  
  lastpcent = pcent
  lastTot = memoryInfo.dwMemoryLoad
  
  GetMemoryInfo = Format(lastpcent)

End Function
Public Function CheckComputerHeal() As String
Dim sRAM
Dim sKQ
sRAM = GetMemoryInfo
If sRAM < 10 Then
    sKQ = "Ma1y Ti1nh D9ang Ra61t Na85ng Va2 Lag, Co1 Le4 Ba5n D9ang Cha5y Ra61t Nhie62u Thu71."
ElseIf sRAM >= 10 And sRAM < 20 Then
    sKQ = "Ma1y Ti1nh D9ang Na85ng Va2 Lag, Co1 Le4 Ba5n D9ang Cha5y Nhie62u Thu71."
ElseIf sRAM >= 20 And sRAM < 30 Then
    sKQ = "Ma1y Ti1nh D9ang Lag"
ElseIf sRAM >= 30 And sRAM < 40 Then
    sKQ = "Ma1y Ti1nh Cha5y O63n D9i5nh, RAM Bi2nh Thu7o72ng."
ElseIf sRAM >= 40 And sRAM < 50 Then
    sKQ = "Ma1y Ti1nh Cha5y Bi2nh Thu7o72ng"
ElseIf sRAM >= 50 And sRAM < 60 Then
    sKQ = "Ma1y Ti1nh Cha5y Nhanh Va2 O63n D9i5nh, Ti2nh Tra5ng To61t."
ElseIf sRAM >= 60 And sRAM < 70 Then
    sKQ = "Ma1y ti1nh D9ang Cha5y Ra61t Nhanh, To61c D9o65 Xu73 Ly1 To61t"
ElseIf sRAM >= 70 And sRAM < 80 Then
    sKQ = "Ma1y Ti1nh D9ang Ra61t To61t"
ElseIf sRAM >= 80 Then
    sKQ = "Ba5n Co1 Mo65t Thanh RAM Tuye65t Vo72i, Ma1y Ti1nh Hoa5t D9o65ng Hoa2n Ha3o D9e61n Tu72ng Chi Tie61t."
End If
    
sKQ = sKQ & " (Free RAM: " & sRAM & " %)"
CheckComputerHeal = sKQ
End Function
Public Sub SetUniText(ByVal hwnd As Long, ByVal sUniText As String)
    DefWindowProcW hwnd, WM_SETTEXT, &H0&, StrPtr(sUniText)
End Sub

Public Function CreateFolderIcon(Ffolder) As String
If FileExists(Ffolder & "\Desktop.ini") = True Then
SetAttr Ffolder & "\Desktop.ini", vbNormal
DeleteFile Ffolder & "\Desktop.ini"
End If
CreateTextFile Ffolder & "\Desktop.ini", "[.ShellClassInfo]" & vbNewLine & _
                                                  "IconFile=" & Ffolder & "\1Click-Protect.ico" & vbNewLine & _
                                                  "IconIndex = 0"
SetAttr Ffolder & "\Desktop.ini", vbReadOnly + vbHidden
End Function
Public Function CheckVirus() As String
Dim strKQ
Dim sCoVirus As Boolean
sCoVirus = False
strKQ = "Ma1y Ti1nh Cu73a Ba5n D9ang Bi5 Nhie64m Virus:"

'C:\WINDOWS\system32\drivers\klif.sys
If FileExists("C:\2fiy.bat") = True Then
strKQ = strKQ & " 2fiy.bat"
sCoVirus = True
End If

If FileExists("C:\WINDOWS\system32\win1ogon.exe") = True Then
strKQ = strKQ & " win1ogon.exe"
sCoVirus = True
End If

If FileExists("C:\WINDOWS\system32\SCVVHSOT.exe") = True Then
strKQ = strKQ & " SCVVHSOT.exe"
sCoVirus = True
End If

If FileExists("C:\WINDOWS\system32\Explorer.sm1") = True Then
 strKQ = strKQ & " sxs.exe"

sCoVirus = True
End If


If FileExists("C:\WINDOWS\megabyte.exe") = True Then
 strKQ = strKQ & " Megabyte.exe"

sCoVirus = True
End If

If FileExists("C:\WINDOWS\System32\Sys\HayTiepTuc.exe") = True Then
 strKQ = strKQ & " TiepTuc.exe"

sCoVirus = True

End If

If FileExists("C:\RESTORE\S-1-5-21-1482476501-1644491937-682003330-1013\Taquito.exe") = True Then
 strKQ = strKQ & " Taquito.exe"

sCoVirus = True

End If

If FileExists("C:\Folder.exe") = True Then
 strKQ = strKQ & " EV-SHUTTLE.exe"

sCoVirus = True

End If


If FileExists("C:\WINDOWS\taskmsg.exe") = True Then
 strKQ = strKQ & " taskmsg.exe"

sCoVirus = True

End If

If FileExists("C:\WINDOWS\system32\scvhosti.exe") = True Then
 strKQ = strKQ & " scvhosti.exe"

sCoVirus = True

End If

If FileExists("C:\Program Files\PCPrivacyCleaner\pcpc.exe") = True Then
 strKQ = strKQ & " pcpc.exe"
sCoVirus = True

End If

If FileExists("C:\WINDOWS\System32\logon.exe") = True Then
 strKQ = strKQ & " algs.exe"

sCoVirus = True

End If

If FileExists("C:\Program Files\AntiMalwareGuard\amg.exe") = True Then
 strKQ = strKQ & " amg.exe"
sCoVirus = True

End If

If FileExists("C:\WINDOWS\system32\IEXPLORER.exe") = True Then
 strKQ = strKQ & " IEXPLORER.exe"
sCoVirus = True

End If

If FileExists("C:\WINDOWS\help\B7C8A6484EE3.exe") = True Then
 strKQ = strKQ & " Shell.exe"
sCoVirus = True

End If


If FileExists("C:\WINDOWS\System32\system.exe") = True Then
 strKQ = strKQ & " Forever.exe"

sCoVirus = True

End If


If FileExists("C:\lbb.com") = True Then
 strKQ = strKQ & "KvoSoft"

sCoVirus = True

End If



If FileExists("C:\zPharaoh.exe") = True Then
 strKQ = strKQ & " zPharaoh.exe"
sCoVirus = True

End If

If FileExists("C:\WINDOWS\phimnguoilon.exe") = True Then
 strKQ = strKQ & " Phimhot.exe"
sCoVirus = True

End If


If FileExists("C:\WINDOWS\Mixa.exe") = True Then
strKQ = strKQ & " Mixa  I.exe"
sCoVirus = True

End If

If sCoVirus = False Then strKQ = "Kho6ng Ti2m Tha61y Virus."
CheckVirus = strKQ
End Function

Public Function CheckTimeUsed() As String
CheckTimeUsed = "Ba5n D9a4 Su73 Du5ng Ma1y Ti1nh D9u7o75c: " & Fix((GetTickCount / 60000) / 60) & " Gio72 " & (Round(GetTickCount / 60000) Mod 60) & " Phu1t"
End Function
Public Function CheckInternet() As String
Dim ret As Long
    ret = InternetGetConnectedStateEx(ret, sConnType, 254, 0)
    If ret = 1 Then
        CheckInternet = "Ma1y Ti1nh Co1 Ke61t No61i Internet."
    Else
       CheckInternet = "Ma1y Ti1nh Kho6ng Ke61t No61i Internet."
    End If
End Function

Private Sub GetComputerInfo()
DoEvents
lblHDH.Caption = GetOS
lblUser.Caption = GetUser
lblComputer.Caption = GetComputer
lblSizeHardDisk.Caption = GetDriveTotal & " GB"
lblSizeRAM.Caption = GetRAMTotal
lblAutorunPro.Caption = AutorunProtect
lblViewAutorun.Caption = AutorunView
lblCheckTask.Caption = CheckTaskManager
lblCheckReg.Caption = CheckRegistry
lblLogOnErr.Caption = CheckLogOnErr
lblComputerHeal.Caption = CheckComputerHeal
lblVirusView.Caption = CheckVirus
lblTimeUse.Caption = CheckTimeUsed
lblInternet.Caption = CheckInternet
DoEvents
End Sub
