VERSION 5.00
Object = "{E8FDD05C-3067-4198-8AEC-1A013A46ABDD}#1.0#0"; "FVUnicodeControl.ocx"
Begin VB.Form frmPlash 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "1Click 1.2.0"
   ClientHeight    =   2475
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   6465
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFlash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrStatus 
      Interval        =   130
      Left            =   720
      Top             =   960
   End
   Begin FVUnicodeControl.FVistaUniLabel lblStatus 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   450
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
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
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   450
      Alignment       =   1
      Caption         =   "Chu7o7ng tri2nh d9ang na5p, vui lo2ng cho72 trong gia6y la1t..."
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
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1440
      Top             =   480
   End
   Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel3 
      Height          =   255
      Left            =   3120
      TabIndex        =   0
      Top             =   1080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      Caption         =   "Version 1.2.0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483631
   End
   Begin FVUnicodeControl.FVistaUniProgressbar Bar1 
      Height          =   225
      Left            =   360
      Top             =   1920
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   397
      Max             =   100
      Value           =   0
      TStyle          =   2
      Min             =   0
      Style           =   1
      Text            =   "Loading..."
      Align           =   1
   End
   Begin FVUnicodeControl.FVistaUniLabel lblMe 
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   0
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   450
      Alignment       =   2
      BackColor       =   -2147483639
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
   Begin VB.Image Image1 
      Height          =   825
      Left            =   2160
      Picture         =   "frmFlash.frx":57E2
      Stretch         =   -1  'True
      Top             =   480
      Width           =   2040
   End
End
Attribute VB_Name = "frmPlash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function DefWindowProcW Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_SETTEXT As Long = &HC



Dim i

Private Sub Timer1_Timer()
DoEvents
Bar1.Value = Bar1.Value + 1
If Bar1.Value = Bar1.Max / 10 Then
DoEvents
Load frmMain
DoEvents
ElseIf Bar1.Value = Bar1.Max Then
DoEvents
frmMain.Show
SetUniText frmMain.hwnd, "1Click - Nhanh ch" & ChrW(243) & "ng - D" & ChrW(7877) & " d" & ChrW(224) & "ng - " & ChrW(272) & ChrW(417) & "n gi" & ChrW(7843) & "n! Phi" & ChrW(234) & "n b" & ChrW(7843) & "n 1.2.0"
If GetString(HKEY_CURRENT_USER, "Software\1Click", "HelpUpdate") = vbNullString Then frmHelpMe.Show
Unload Me
DoEvents
End If
DoEvents
End Sub
Public Sub SetUniText(ByVal hwnd As Long, ByVal sUniText As String)
    DefWindowProcW hwnd, WM_SETTEXT, &H0&, StrPtr(sUniText)
End Sub

Private Sub tmrStatus_Timer()
Dim sText(10) As String

    sText(1) = "Ca65p Nha65t Phie6n Ba3n..."
    sText(2) = "Ca2i D9a85t OCX..."
    sText(3) = "La61y Tho6ng Tin Ma1y Ti1nh..."
    sText(4) = "Kie63m Tra Ca1c Chu71c Na8ng Bi5 Kho1a..."
    sText(5) = "Kie63m Tra Ca1c Lo64i Co1 Trong Ma1y..."
    sText(6) = "Kie63m Tra Danh Sa1ch Virus..."
    sText(7) = "Kie63m Tra Autorun..."
    sText(8) = "Kie63m Tra Lo64i Registry..."
    sText(9) = "Kie63m Tra Ca2i D9a85t..."
    sText(10) = "Hoa2n Ta61t...!"

    i = i + 1
If i > 9 Then
        tmrStatus.Enabled = False
    End If

    lblStatus.Caption = sText(i)

End Sub
