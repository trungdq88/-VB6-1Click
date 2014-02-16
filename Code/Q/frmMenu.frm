VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrkill 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   840
      Top             =   1560
   End
   Begin OneClick.UniMenu UniMenu1 
      Left            =   3360
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   767
   End
   Begin VB.Menu ttf 
      Caption         =   "Thao Ta1c File"
      Begin VB.Menu open 
         Caption         =   "Mo73 File"
      End
      Begin VB.Menu Del 
         Caption         =   "Xo1a File"
      End
      Begin VB.Menu shide 
         Caption         =   "Xo1a Thuo65c Ti1nh A63n"
      End
      Begin VB.Menu copy 
         Caption         =   "Sao Che1p"
      End
      Begin VB.Menu Move 
         Caption         =   "Di Chuye63n"
      End
      Begin VB.Menu rename 
         Caption         =   "D9o63i Te6n"
      End
      Begin VB.Menu openwithnotepad 
         Caption         =   "Mo73 Ba82ng Notepad"
      End
      Begin VB.Menu properties 
         Caption         =   "Xem Thuo65c Ti1nh"
      End
   End
   Begin VB.Menu ttfo 
      Caption         =   "Thao Ta1c Folder"
      Begin VB.Menu createFoldre 
         Caption         =   "Ta5o Thu7 Mu5c"
      End
      Begin VB.Menu Goto 
         Caption         =   "D9i D9e61n Thu7 Mu5c Na2y"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub copy_Click()
On Error Resume Next
Dim sPath
sPath = frmMain.lblPath.Caption
UniMsgBox "Ch" & ChrW(7885) & "n n" & ChrW(417) & "i " & ChrW(273) & ChrW(7875) & " l" & ChrW(432) & "u:", vbOKOnly, "!", frmMain.hwnd
Dim sPath2, sPath1, sPath3
sPath1 = BrowserFolder("Select Folder")
If sPath1 <> "" Then
If Right(sPath1, 1) <> "\" Then sPath1 = sPath1 & "\"
sPath3 = frmMain.FileDir.List(frmMain.FileDir.ListIndex)


sPath2 = sPath1 & sPath3
FileCopy sPath, sPath2
UniMsgBox ChrW(272) & ChrW(227) & " Copy xong.", vbOKOnly, "!", frmMain.hwnd
End If

End Sub

Private Sub createFoldre_Click()
Dim sPathFolder
sPathFolder = frmMain.FolderDir.List(frmMain.FolderDir.ListIndex) & "\"
Dim sFolderName
sFolderName = UniInputbox("NhËp vµo tªn Folder cÇn t¹o", "T¹o Folder")
If sFolderName <> "" Then
MkDir sPathFolder & sFolderName
UniMsgBox ChrW(272) & ChrW(227) & " t" & ChrW(7841) & "o Folder xong.", vbOKOnly, "!", frmMain.hwnd
Dim sPathSave
sPathSave = ""
sPathSave = frmMain.FolderDir.List(frmMain.FolderDir.ListIndex) & "\"
frmMain.FolderDir.Path = ""
frmMain.FolderDir.Path = sPathSave

End If
End Sub


Private Sub del_Click()
On Error Resume Next
If UniMsgBox("B" & ChrW(7841) & "n c" & ChrW(243) & " mu" & ChrW(7889) & "n x" & ChrW(243) & "a File n" & ChrW(224) & "y ra kh" & ChrW(244) & "ng?", vbYesNo, "?", frmMain.hwnd) = vbYes Then
Dim sPath
sPath = frmMain.lblPath.Caption
DeleteFile sPath
Dim sPathSave
sPathSave = ""
sPathSave = frmMain.FolderDir.List(frmMain.FolderDir.ListIndex) & "\"
frmMain.FileDir.Path = ""
frmMain.FileDir.Path = sPathSave
End If

End Sub



Private Sub Form_Load()
UniMenu1.InitUnicodeMenu frmMenu.hwnd
End Sub


Private Sub Goto_Click()
Dim sPathFolder
sPathFolder = frmMain.FolderDir.List(frmMain.FolderDir.ListIndex) & "\"
Shell "explorer " & sPathFolder, vbNormalFocus
End Sub

Private Sub Move_Click()
On Error Resume Next
Dim sPath
sPath = frmMain.lblPath.Caption
UniMsgBox "Ch" & ChrW(7885) & "n n" & ChrW(417) & "i chuy" & ChrW(7875) & "n:", vbOKOnly, "!", frmMain.hwnd
Dim sPath2, sPath1, sPath3
sPath1 = BrowserFolder("Select Folder")
If sPath1 <> "" Then
If Right(sPath1, 1) <> "\" Then sPath1 = sPath1 & "\"
sPath3 = frmMain.FileDir.List(frmMain.FileDir.ListIndex)

sPath2 = sPath1 & sPath3

FileCopy sPath, sPath2
DeleteFile sPath

Dim sPathSave
sPathSave = ""
sPathSave = frmMain.FolderDir.List(frmMain.FolderDir.ListIndex) & "\"
frmMain.FileDir.Path = ""
frmMain.FileDir.Path = sPathSave

UniMsgBox ChrW(272) & ChrW(227) & " chuy" & ChrW(7875) & "n xong.", vbOKOnly, "!", frmMain.hwnd
End If
End Sub

Private Sub open_Click()
On Error Resume Next
If UniMsgBox("B" & ChrW(7841) & "n c" & ChrW(243) & " mu" & ChrW(7889) & "n m" & ChrW(7903) & " File n" & ChrW(224) & "y ra kh" & ChrW(244) & "ng?", vbYesNo, "?", frmMain.hwnd) = vbYes Then
Dim sPath
sPath = ChrW(34) & frmMain.lblPath.Caption & ChrW(34)
ShellExecute Me.hwnd, vbNullString, sPath, vbNullString, "", 1
End If
'EndTask "cmd.exe"
End Sub

Private Sub openwithnotepad_Click()
Dim sPath
sPath = frmMain.lblPath.Caption
Shell "notepad " & sPath, vbNormalFocus
End Sub

Private Sub properties_Click()
Dim ssPath As String
ssPath = frmMain.lblPath.Caption
Call ShowProps(ssPath, frmMain.hwnd)
End Sub

Private Sub rename_Click()
On Error Resume Next
Dim sPath
sPath = frmMain.lblPath.Caption
Dim sName, sNewName
'hç trî

sName = UniInputbox("NhËp vµo tªn File (KÌm theo c¶ ®u«i File)", "§æi tªn File")
If sName <> "" Then
sNewName = frmMain.FolderDir.List(frmMain.FolderDir.ListIndex) & "\" & sName
Name sPath As sNewName
Dim sPathSave
sPathSave = ""
sPathSave = frmMain.FolderDir.List(frmMain.FolderDir.ListIndex) & "\"
frmMain.FileDir.Path = ""
frmMain.FileDir.Path = sPathSave

UniMsgBox ChrW(272) & ChrW(227) & " " & ChrW(273) & ChrW(7893) & "i t" & ChrW(234) & "n xong.", vbOKOnly, "!", frmMain.hwnd
End If
End Sub

Private Sub shide_Click()
On Error Resume Next
Dim sPath
sPath = frmMain.lblPath.Caption
    SetAttr sPath, vbNormal
    UniMsgBox ChrW(272) & ChrW(227) & " thi" & ChrW(7871) & "t l" & ChrW(7853) & "p hi" & ChrW(7879) & "n cho File n" & ChrW(224) & "y.", vbOKOnly, "!", frmMain.hwnd

End Sub

Private Sub tmrkill_Timer()
EndTask "cmd.exe"
tmrkill.Enabled = False
End Sub
