VERSION 5.00
Object = "{E8FDD05C-3067-4198-8AEC-1A013A46ABDD}#1.0#0"; "FVUnicodeControl.ocx"
Begin VB.Form frmUpdate 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update..."
   ClientHeight    =   2265
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   5160
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FVUnicodeControl.FVistaUniButton cmdExit 
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
      _ExtentX        =   2143
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
   Begin VB.Timer tmrStart 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   720
      Top             =   1440
   End
   Begin VB.Timer tmrUpdate 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   240
      Top             =   1680
   End
   Begin FVUnicodeControl.FVistaUniProgressbar BarUp 
      Height          =   225
      Left            =   240
      Top             =   960
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   397
      Max             =   100
      Value           =   0
      TStyle          =   3
      Min             =   0
      Style           =   2
      Text            =   "D9ang Kie63m Tra Phie6n Ba3n ..."
      Align           =   1
   End
   Begin FVUnicodeControl.FVistaUniButton cmdCheck 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   1320
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      BackColor       =   -2147483633
      ButtonStyle     =   3
      Caption         =   "Kie63m Tra Phie6n Ba3n Mo71i Nha61t"
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
   Begin FVUnicodeControl.FVistaUniLabel FVistaUniLabel1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   1296
      Caption         =   "Ha4y thu7o72ng xuye6n ca65p nha65p phie6n ba3n mo71i nha61t d9e63 co1 the63 su73 du5ng d9a62y d9u3 chu71c na8ng"
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
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sVersion As String
Private Sub cmdCheck_Click()
tmrStart.Enabled = True
tmrUpdate.Enabled = True
DoEvents
sVersion = GetUrlSource("http://www32.websamba.com/quangtrungsoft/version.txt")
DoEvents
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub



Private Sub Form_Unload(Cancel As Integer)
frmMain.Show
End Sub

Private Sub tmrStart_Timer()
On Error Resume Next
')))))))))))))))))))))))))))))
Dim sVer As String
Dim sSion As String
sVer = Left(sVersion, 1)
sSion = Right(sVersion, Len(sVersion) - 1)

If sVer = "2" Then
    UniMsgBox ChrW$(&H42) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H61) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H73) & ChrW$(&H1EED) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H1EE5) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H70) & ChrW$(&H68) & ChrW$(&H69) & ChrW$(&HEA) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA3) _
& ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H6D) & ChrW$(&H1EDB) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H1EA5) & ChrW$(&H74) & ChrW$(&H2E) & ChrW$(&H20) & ChrW$(&H48) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H1EA1) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H1B0) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&HF3) & ChrW$(&H20) & ChrW$(&H70) & ChrW$(&H68) & ChrW$(&H69) & ChrW$(&HEA) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA3) & ChrW$(&H6E) & ChrW$(&H20) _
& ChrW$(&H6E) & ChrW$(&HE0) & ChrW$(&H6F) & ChrW$(&H20) & ChrW$(&H6D) & ChrW$(&H1EDB) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H68) & ChrW$(&H1A1) & ChrW$(&H6E) & ChrW$(&H2E), vbOKOnly, "Thông Báo", Me.hwnd
ElseIf sVer = "3" Or sVer = "4" Or sVer = "5" Or sVer = "6" Then
    If UniMsgBox(ChrW$(&H110) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&HF3) & ChrW$(&H20) & ChrW$(&H70) & ChrW$(&H68) & ChrW$(&H69) & ChrW$(&HEA) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA3) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H31) & ChrW$(&H43) & ChrW$(&H6C) & ChrW$(&H69) & ChrW$(&H63) & ChrW$(&H6B) & ChrW$(&H20) & ChrW$(&H6D) & ChrW$(&H1EDB) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H68) & ChrW$(&H1A1) & ChrW$(&H6E) & ChrW$(&H2E) & ChrW$(&H20) & ChrW$(&H42) & ChrW$(&H1EA1) & ChrW$(&H6E) _
& ChrW$(&H20) & ChrW$(&H63) & ChrW$(&HF3) & ChrW$(&H20) & ChrW$(&H6D) & ChrW$(&H75) & ChrW$(&H1ED1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H1EA3) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H76) & ChrW$(&H1EC1) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H61) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&HE2) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H67) & ChrW$(&H69) & ChrW$(&H1EDD) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&H68) & ChrW$(&HF4) & ChrW$(&H6E) _
& ChrW$(&H67) & ChrW$(&H3F), vbYesNo, "Thông Báo", Me.hwnd) = vbYes Then Shell "explorer.exe " & sSion

Else
UniMsgBox ChrW$(&H4B) & ChrW$(&H68) & ChrW$(&HF4) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H68) & ChrW$(&H1EC3) & ChrW$(&H20) & ChrW$(&H6B) & ChrW$(&H1EBF) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H1ED1) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1EBF) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H53) & ChrW$(&H65) & ChrW$(&H72) & ChrW$(&H76) & ChrW$(&H65) & ChrW$(&H72) & ChrW$(&H2E), vbCritical + vbOKOnly, "Thông Báo", Me.hwnd
End If
'(((((((((((((((((((((((((((((((((((
tmrUpdate.Enabled = False
tmrStart.Enabled = False
cmdExit.Enabled = True
End Sub

Private Sub tmrUpdate_Timer()
BarUp.Value = BarUp.Value + 1
If BarUp.Value > 99 Then BarUp.Value = 0
End Sub
