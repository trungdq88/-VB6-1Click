VERSION 5.00
Object = "{E8FDD05C-3067-4198-8AEC-1A013A46ABDD}#1.0#0"; "FVUnicodeControl.ocx"
Begin VB.Form frmWait 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   0  'None
   Caption         =   "Please Wait..."
   ClientHeight    =   825
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   825
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   360
      Top             =   600
   End
   Begin FVUnicodeControl.FVistaUniProgressbar P1 
      Height          =   225
      Left            =   120
      Top             =   240
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   397
      Max             =   100
      Value           =   0
      TStyle          =   3
      Min             =   0
      Style           =   2
      Text            =   "Chu7o7ng Tri2nh D9ang Xu73 Ly1, Vui Lo2ng Cho72 Trong Gia6y La1t..."
      Align           =   1
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Timer1_Timer()
P1.Value = P1.Value + 2
If P1.Value > P1.Max - 4 Then P1.Value = 0
End Sub
