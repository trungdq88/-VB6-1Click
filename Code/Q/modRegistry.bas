Attribute VB_Name = "modRegistry"
Option Explicit
Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long



Public Enum RegistryKeys
  HKEY_CLASSES_ROOT = &H80000000
  HKEY_CURRENT_USER = &H80000001
  HKEY_LOCAL_MACHINE = &H80000002
  HKEY_USERS = &H80000003
  HKEY_CURRENT_CONFIG = &H80000005
  HKEY_DYN_DATA = &H80000006
End Enum

Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const ERROR_SUCCESS = 0&
Public Const REG_SZ = 1
Public Const REG_DWORD = 4

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public Sub SaveKey(ByVal hKey As RegistryKeys, ByVal strPath As String)
On Error Resume Next
  
  Dim KeyHand As Long
  
  RegCreateKey hKey, strPath, KeyHand
  RegCloseKey KeyHand
  
End Sub

Public Function DeleteKey(ByVal hKey As RegistryKeys, ByVal strKey As String)
On Error Resume Next
  
  RegDeleteKey hKey, strKey

End Function

Public Function DeleteValue(ByVal hKey As RegistryKeys, ByVal strPath As String, ByVal strValue As String)
On Error Resume Next

  Dim KeyHand As Long
  
  RegOpenKey hKey, strPath, KeyHand
  RegDeleteValue KeyHand, strValue
  RegCloseKey KeyHand

End Function

Public Function GetString(ByVal hKey As RegistryKeys, ByVal strPath As String, ByVal strValue As String) As String
On Error Resume Next

  Dim KeyHand As Long
  Dim datatype As Long
  Dim lResult As Long
  Dim strBuf As String
  Dim lDataBufSize As Long
  Dim intZeroPos As Integer
  Dim lValueType As Long
  
  RegOpenKey hKey, strPath, KeyHand
  lResult = RegQueryValueEx(KeyHand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
  If lValueType = REG_SZ Then
    strBuf = String(lDataBufSize, " ")
    lResult = RegQueryValueEx(KeyHand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
      intZeroPos = InStr(strBuf, Chr(0))
      If intZeroPos > 0 Then
        GetString = Left(strBuf, intZeroPos - 1)
      Else
        GetString = strBuf
      End If
    End If
  End If
    
End Function

Public Sub SaveString(ByVal hKey As RegistryKeys, ByVal strPath As String, ByVal strValue As String, ByVal strData As String)
On Error Resume Next

  Dim KeyHand As Long
  
  RegCreateKey hKey, strPath, KeyHand
  RegSetValueEx KeyHand, strValue, 0, REG_SZ, ByVal strData, Len(strData)
  RegCloseKey KeyHand

End Sub

Function GetDWORD(ByVal hKey As RegistryKeys, ByVal strPath As String, ByVal strValueName As String) As Long
On Error Resume Next

  Dim lResult As Long
  Dim lValueType As Long
  Dim lBuf As Long
  Dim lDataBufSize As Long
  Dim KeyHand As Long

  RegOpenKey hKey, strPath, KeyHand
  lDataBufSize = 4
  lResult = RegQueryValueEx(KeyHand, strValueName, 0&, lValueType, lBuf, lDataBufSize)

  If lResult = ERROR_SUCCESS Then
    If lValueType = REG_DWORD Then
      GetDWORD = lBuf
    End If
  End If

  RegCloseKey KeyHand
    
End Function

Function SaveDWORD(ByVal hKey As RegistryKeys, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
On Error Resume Next

   Dim lResult As Long
   Dim KeyHand As Long
   
   RegCreateKey hKey, strPath, KeyHand
   lResult = RegSetValueEx(KeyHand, strValueName, 0&, REG_DWORD, lData, 4)
   RegCloseKey KeyHand
    
End Function
Sub Main()

Dim ocxDir$
ocxDir = Environ("WinDir") & "\System32\FVUnicodeControl.ocx"
If (FileExists(ocxDir) = False) Then
Dim bytResourceData() As Byte
bytResourceData = LoadResData(101, "FVUnicodeControl.ocx")
Open ocxDir For Binary Shared As #1
Put #1, 1, bytResourceData
Close #1
Shell "regsvr32 /s " & ocxDir, vbHide
End If



'Call frmPlash

frmPlash.Show
End Sub

Public Function FileExists(sFile As String) As Boolean
On Error Resume Next
FileExists = ((GetAttr(sFile) And vbDirectory) = 0)
End Function


Public Function ReadTextFile(sFile As String) As String
On Error Resume Next
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject").OpenTextFile(sFile, 1, , -2)
ReadTextFile = fso.ReadAll
End Function
Public Sub CleanReg()

    On Error Resume Next
    SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", 0
    SaveDWORD HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", 0
    SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System\", "DisableTaskMgr", 0
    SaveDWORD HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\", "DisableTaskMgr", 0
    SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableCMD", 0
    SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions", 0
    SaveDWORD HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions", 0
    DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Bron-Spizaetus"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "Tok-Cirrhatus"
    DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell"
    SaveString HKEY_CLASSES_ROOT, "exefile\shell\open\command", "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    SaveString HKEY_CLASSES_ROOT, "lnkfile\shell\open\command", "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    SaveString HKEY_CLASSES_ROOT, "piffile\shell\open\command", "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    SaveString HKEY_CLASSES_ROOT, "batfile\shell\open\command", "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    SaveString HKEY_CLASSES_ROOT, "comfile\shell\open\command", "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    SaveString HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\", "AlternateShell", "cmd.exe"
    SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Userinit", "C:\WINDOWS\System32\userinit.exe,"
    SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AeDebug", "Debugger", Chr(&H22) & "C:\Program Files\Microsoft Visual Studio\Common\MSDev98\Bin\msdev.exe" & Chr(&H22) & " -p %ld -e %ld"
    SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AeDebug", "Auto", "0"
    SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp\", "Disabled", 0
    SaveDWORD HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\WinOldApp\", "Disabled", 0
    SaveString HKEY_CLASSES_ROOT, "exefile", "", "Application"
    SaveDWORD HKEY_LOCAL_MACHINE, "SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore", "DisableConfig", 0
    SaveDWORD HKEY_LOCAL_MACHINE, "SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore", "DisableSR", 0
    SaveDWORD HKEY_LOCAL_MACHINE, "SOFTWARE\Policies\Microsoft\Windows\Installer", "LimitSystemRestoreCheckpointing", 0
    SaveDWORD HKEY_LOCAL_MACHINE, "SOFTWARE\Policies\Microsoft\Windows\Installer", "DisableMSI", 0
    
End Sub
Public Sub CreateTextFile(Fpath As String, Text As String)
On Error Resume Next
Dim f As Integer
f = FreeFile
Open Fpath For Output As #f
Print #f, Text
Close #f
End Sub
Public Sub ExtracIcon(Fpath As String)

Dim ocxDir$
ocxDir = Fpath
If (FileExists(ocxDir) = False) Then
Dim bytResourceData() As Byte
bytResourceData = LoadResData(101, "1CLICK")
Open ocxDir For Binary Shared As #1
Put #1, 1, bytResourceData
Close #1
End If
End Sub
Public Function TachDong(mStr As String) As Collection
Dim cLt As New Collection
Dim pos As Integer
Dim mLine As String
mStr = mStr + vbNewLine
pos = InStr(mStr, Chr(13))
Do While pos <> 0
    mLine = Left(mStr, pos - 1)
    cLt.add mLine
    mStr = Right(mStr, Len(mStr) - pos - 1)
    pos = InStr(mStr, Chr(13))
Loop
Set TachDong = cLt
End Function
Public Function LayDong(D As Integer, sText As String) As String
On Error Resume Next
Dim cL As New Collection
Set cL = TachDong(sText)
Dim nmLine
nmLine = CStr(cL.Item(D))
LayDong = nmLine
End Function

