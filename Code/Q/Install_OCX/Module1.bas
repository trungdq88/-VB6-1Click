Attribute VB_Name = "Module1"
Sub Main()

Dim ocxDir$
'Get OCX Directory
ocxDir = Environ("WinDir") & "\System32\FVUnicodeControl.ocx"
If (FileExists(ocxDir) = False) Then
'Get OCX on Resource Data
Dim bytResourceData() As Byte
bytResourceData = LoadResData(101, "FVUnicodeControl.ocx")
'Save OCX as Directory
Open ocxDir For Binary Shared As #1
Put #1, 1, bytResourceData
Close #1

'Reg OCX
Shell "regsvr32 /s " & ocxDir, vbHide
End If


'shdocvw.dll

ocxDir = Environ("WinDir") & "\System32\Comdlg32.ocx"
If (FileExists(ocxDir) = False) Then
'Get OCX on Resource Data
bytResourceData = LoadResData(101, "Comdlg32.ocx")
'Save OCX as Directory
Open ocxDir For Binary Shared As #1
Put #1, 1, bytResourceData
Close #1

'Reg OCX
Shell "regsvr32 /s " & ocxDir, vbHide
End If



ocxDir = Environ("WinDir") & "\System32\shdocvw.dll"
If (FileExists(ocxDir) = False) Then
'Get OCX on Resource Data
bytResourceData = LoadResData(101, "shdocvw.dll")
'Save OCX as Directory
Open ocxDir For Binary Shared As #1
Put #1, 1, bytResourceData
Close #1

'Reg OCX
Shell "regsvr32 /s " & ocxDir, vbHide
End If
frmInstall.Show
End Sub



Public Function FileExists(sFile As String) As Boolean
On Error Resume Next
FileExists = ((GetAttr(sFile) And vbDirectory) = 0)
End Function


