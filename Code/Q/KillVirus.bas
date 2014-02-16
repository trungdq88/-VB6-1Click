Attribute VB_Name = "KillVirus"
Public Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal Handle As Long) As Long
Public Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Public Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Public Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Public Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

'khai báo ki?u d? li?u c?n dùng
Public Type PROCESSENTRY32
   dwSize As Long
   cntUsage As Long
   th32ProcessID As Long ' This process
   th32DefaultHeapID As Long
   th32ModuleID As Long
' Associated exe
   cntThreads As Long
   th32ParentProcessID As Long
' This process's parent process
   pcPriClassBase As Long
' Base priority of process threads
   dwFlags As Long
   szexeFile As String * 260 ' MAX_PATH
End Type

Public Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
'1 = Windows 95, 2 = Windows NT
szCSDVersion As String * 128
End Type

Public Const PROCESS_QUERY_INFORMATION = 1024
Public Const PROCESS_VM_READ = 16
Public Const PROCESS_TERMINATE = 1
Public Const MAX_PATH = 260
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const SYNCHRONIZE = &H100000
Public Const PROCESS_ALL_ACCESS = &H1F0FFF
Public Const TH32CS_SNAPPROCESS = &H2&
Public Const hNull = 0

Declare Function ProcessFirst Lib "kernel32.dll" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Declare Function ProcessNext Lib "kernel32.dll" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Declare Function CreateToolhelpSnapshot Lib "kernel32.dll" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long

'-------------------------------------------------------
Public Function FindProcessID(NameProcess As String) As Long
Const PROCESS_ALL_ACCESS = &H1F0FFF
Const TH32CS_SNAPPROCESS As Long = 2&
Dim uProcess  As PROCESSENTRY32
Dim RProcessFound As Long
Dim hSnapshot As Long
Dim SzExename As String
Dim i As Integer
        
       If NameProcess <> "" Then
 
          uProcess.dwSize = Len(uProcess)
          hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
          RProcessFound = ProcessFirst(hSnapshot, uProcess)
  
          Do
            i = InStr(1, uProcess.szexeFile, Chr(0))
            SzExename = LCase$(Left$(uProcess.szexeFile, i - 1))
        
            If Right$(SzExename, Len(NameProcess)) = LCase$(NameProcess) Then
               FindProcessID = uProcess.th32ProcessID
               Exit Do
            End If
            RProcessFound = ProcessNext(hSnapshot, uProcess)
          Loop While RProcessFound
          Call CloseHandle(hSnapshot)
       End If
 
End Function


Function EndTask(Proccess As String)
 Dim hProcess As Long
   Dim RetVal As Long

   hProcess = OpenProcess(SYNCHRONIZE Or PROCESS_TERMINATE, 0, CLng(FindProcessID(Proccess)))
If hProcess <> 0 Then

   RetVal = TerminateProcess(hProcess, 0)

End If
End Function


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Sub KillMixa()
frmWait.Show
On Error Resume Next
EndTask "mixa.exe" 'taskkill /f /im mixa.exe
EndTask "mixa_I.exe" 'taskkill /f /im mixa_I.exe
EndTask "systemio.exe" 'taskkill /f /im systemio.exe

SetAttr "C:\WINDOWS\Mixa.exe", vbNormal 'attrib -h -s -r c:\windows\mixa.exe
SetAttr "c:\windows\system32\systemio.exe", vbNormal
SetAttr "C:\mixa_I.exe", vbNormal 'attrib -h -s -r c:\mixa_I.exe
SetAttr "D:\mixa_I.exe", vbNormal 'attrib -h -s -r d:\mixa_I.exe
SetAttr "E:\mixa_I.exe", vbNormal 'attrib -h -s -r e:\mixa_I.exe
SetAttr "C:\autorun.inf", vbNormal
SetAttr "D:\autorun.inf", vbNormal
SetAttr "E:\autorun.inf", vbNormal

DeleteFile "C:\WINDOWS\Mixa.exe"
DeleteFile "C:\Mixa.exe"
DeleteFile "D:\Mixa.exe"
DeleteFile "E:\Mixa.exe"
DeleteFile "C:\Autorun.inf"
DeleteFile "D:\Autorun.inf"
DeleteFile "E:\Autorun.inf"
DeleteFile "C:\windows\system32\systemio.exe"

DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Virus"
DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Userinit"
SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Userinit", "C:\WINDOWS\system32\userinit.exe"
SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", "explorer.exe"

DeleteValue HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Virus"
DeleteValue HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Userinit"
SaveString HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Userinit", "C:\WINDOWS\system32\userinit.exe"
SaveString HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", "explorer.exe"

EndTask "explorer.exe"
Unload frmWait
UniMsgBox ChrW$(&H110) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H44) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H58) & ChrW$(&H6F) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H21), vbOKOnly, "Thông Báo", frmMain.hWnd
End Sub

Sub KillPhimHot()
frmWait.Show
On Error Resume Next
'Secret.exe
EndTask "explorer.exe"
EndTask "phimhot.exe"
EndTask "Secret.exe"

DeleteFile "C:\WINDOWS\phimnguoilon.exe"
DeleteFile "C:\WINDOWS\Secret.exe"

DeleteFile "C:\Secret.exe"
DeleteFile "D:\Secret.exe"
DeleteFile "E:\Secret.exe"

DeleteFile "C:\Phimhot.exe"
DeleteFile "D:\Phimhot.exe"
DeleteFile "E:\Phimhot.exe"

DeleteFile "C:\phimnguoilon.exe"
DeleteFile "D:\phimnguoilon.exe"
DeleteFile "E:\phimnguoilon.exe"

DeleteFile "C:\autorun.inf"
DeleteFile "D:\autorun.inf"
DeleteFile "E:\autorun.inf"

SaveDWORD HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\wscsvc", "AutorunsDisabled", 1
SaveDWORD HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\wscsvc", "Start", 4
SaveDWORD HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\wuauserv", "AutorunsDisabled", 1
SaveDWORD HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Services\wuauserv", "Start", 4
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Search Assistant", "SocialUI", 0
SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Userinit", "C:\WINDOWS\system32\userinit.exe"
SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", "explorer.exe"
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", 0
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr", 0
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions", 0
DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Folder\Hidden\ShowAll", "CheckedValue"
SaveDWORD HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Folder\Hidden\ShowAll", "CheckedValue", 1
EndTask "explorer.exe"
Unload frmWait

UniMsgBox ChrW$(&H110) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H44) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H58) & ChrW$(&H6F) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H21), vbOKOnly, "Thông Báo", frmMain.hWnd

End Sub

Sub KillImages()
frmWait.Show
On Error Resume Next
EndTask "explorer.exe"
EndTask "ctfmon.exe"
EndTask "system.exe"
EndTask "userinit.exe"

DeleteFile "C:\WINDOWS\userinit.exe"
DeleteFile "C:\WINDOWS\system32\system.exe"
DeleteFile "C:\WINDOWS\system volume information.exe"

SetAttr "C:\autorun.inf", vbNormal
SetAttr "D:\autorun.inf", vbNormal
SetAttr "E:\autorun.inf", vbNormal
SetAttr "C:\Images.exe", vbNormal
SetAttr "D:\Images.exe", vbNormal
SetAttr "E:\Images.exe", vbNormal

DeleteFile "C:\Autorun.inf"
DeleteFile "D:\Autorun.inf"
DeleteFile "E:\Autorun.inf"
DeleteFile "C:\Images.exe"
DeleteFile "D:\Images.exe"
DeleteFile "E:\Images.exe"

SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Userinit", "C:\WINDOWS\system32\userinit.exe"
SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", "explorer.exe"
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", 0
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr", 0
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions", 0
DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Folder\Hidden\ShowAll", "CheckedValue"
SaveDWORD HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Folder\Hidden\ShowAll", "CheckedValue", 1
Shell "explorer.exe"
Unload frmWait

UniMsgBox ChrW$(&H110) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H44) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H58) & ChrW$(&H6F) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H21), vbOKOnly, "Thông Báo", frmMain.hWnd

End Sub

Sub KillzPharaoh()
On Error Resume Next

EndTask "cmd.exe"
EndTask "zPharaoh.exe"
EndTask "explorer.exe"

Kill "C:\*.taz"
Kill "D:\*.taz"
Kill "E:\*.taz"

DeleteFile "C:\zPharaoh.exe"
DeleteFile "D:\zPharaoh.exe"
DeleteFile "E:\zPharaoh.exe"

RmDir "C:\zPharaoh"
RmDir "D:\zPharaoh"
RmDir "E:\zPharaoh"

Kill "C:\Documents and Settings\tazebama*"
Kill "D:\Documents and Settings\tazebama*"
Kill "C:\Documents and Settings\hook*"
Kill "D:\Documents and Settings\hook*"

Kill "C:\*.???.exe"
Kill "D:\*.???.exe"
Kill "E:\*.???.exe"

DeleteFile "C:\autorun.inf"
DeleteFile "D:\autorun.inf"
DeleteFile "E:\autorun.inf"

Shell "explorer.exe"
UniMsgBox ChrW$(&H110) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H44) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H58) & ChrW$(&H6F) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H21), vbOKOnly, "Thông Báo", frmMain.hWnd

End Sub
Sub KillKvoSoft()
On Error Resume Next
EndTask "explorer.exe"
Kill "C:\WINDOWS\System32\kvosoft.*"
Kill "C:\WINDOWS\System32\dsetwem?.*"
'Kvosoft.exe
DeleteFile "C:\autorun.inf"
DeleteFile "D:\autorun.inf"
DeleteFile "E:\autorun.inf"
DeleteFile "C:\Kvosoft.exe"
DeleteFile "D:\Kvosoft.exe"
DeleteFile "E:\Kvosoft.exe"

SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Userinit", "C:\WINDOWS\system32\userinit.exe"
SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", "explorer.exe"

DeleteFile "C:\lbb.com"
DeleteFile "D:\lbb.com"
DeleteFile "E:\lbb.com"

EndTask "explorer.exe"
UniMsgBox ChrW$(&H110) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H44) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H58) & ChrW$(&H6F) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H21), vbOKOnly, "Thông Báo", frmMain.hWnd

End Sub
Sub KillForever()
On Error Resume Next
EndTask "explorer.exe"
EndTask "system.exe"
EndTask "userinit.exe"
SetAttr "C:\forever.exe", vbNormal
SetAttr "D:\forever.exe", vbNormal
SetAttr "E:\forever.exe", vbNormal
SetAttr "C:\Secret.exe", vbNormal
SetAttr "D:\Secret.exe", vbNormal
SetAttr "E:\Secret.exe", vbNormal

DeleteFile "C:\forever.exe"
DeleteFile "D:\forever.exe"
DeleteFile "E:\forever.exe"
DeleteFile "C:\Secret.exe"
DeleteFile "D:\Secret.exe"
DeleteFile "E:\Secret.exe"

SetAttr "C:\WINDOWS\userinit.exe", vbNormal
SetAttr "C:\WINDOWS\System32\system.exe", vbNormal
SetAttr "C:\autorun.inf", vbNormal
SetAttr "D:\autorun.inf", vbNormal
SetAttr "E:\autorun.inf", vbNormal

DeleteFile "C:\WINDOWS\userinit.exe"
DeleteFile "C:\WINDOWS\System32\system.exe"
DeleteFile "C:\autorun.inf"
DeleteFile "D:\autorun.inf"
DeleteFile "E:\autorun.inf"

SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Userinit", "C:\WINDOWS\system32\userinit.exe"
SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", "explorer.exe"

EndTask "explorer.exe"

UniMsgBox ChrW$(&H110) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H44) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H58) & ChrW$(&H6F) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H21), vbOKOnly, "Thông Báo", frmMain.hWnd

End Sub
Sub KillShell()
On Error Resume Next
EndTask "explorer.exe"
SetAttr "C:\WINDOWS\help\B7C8A6484EE3.exe", vbNormal
SetAttr "C:\WINDOWS\help\B7C8A6484EE3.dll", vbNormal
SetAttr "C:\WINDOWS\help\Autorun.inf", vbNormal
SetAttr "C:\WINDOWS\1.bat", vbNormal
SetAttr "C:\autorun.inf", vbNormal
SetAttr "C:\Shell.exe", vbNormal
SetAttr "D:\autorun.inf", vbNormal
SetAttr "D:\Shell.exe", vbNormal
SetAttr "E:\autorun.inf", vbNormal
SetAttr "E:\Shell.exe", vbNormal
SetAttr "F:\autorun.inf", vbNormal
SetAttr "F:\Shell.exe", vbNormal

DeleteFile "C:\WINDOWS\help\B7C8A6484EE3.exe"
DeleteFile "C:\WINDOWS\help\B7C8A6484EE3.dll"
DeleteFile "C:\WINDOWS\help\Autorun.inf"
DeleteFile "C:\WINDOWS\1.bat"
DeleteFile "C:\autorun.inf"
DeleteFile "C:\Shell.exe"
DeleteFile "D:\autorun.inf"
DeleteFile "D:\Shell.exe"
DeleteFile "E:\autorun.inf"
DeleteFile "E:\Shell.exe"
DeleteFile "F:\autorun.inf"
DeleteFile "F:\Shell.exe"

DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\CLSID\{6FC2B704-28A3-464F-AEA2-034E1107B0C4}"
DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellExecuteHooks", "{6FC2B704-28A3-464F-AEA2-034E1107B0C4}"

SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Userinit", "C:\WINDOWS\system32\userinit.exe"
SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", "explorer.exe"

EndTask "explorer.exe"

UniMsgBox ChrW$(&H110) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H44) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H58) & ChrW$(&H6F) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H21), vbOKOnly, "Thông Báo", frmMain.hWnd
End Sub
Sub KillALGS()
On Error Resume Next
EndTask "logon.exe"
SetAttr "C:\WINDOWS\System32\logon.exe", vbNormal
DeleteFile "C:\WINDOWS\System32\logon.exe"
DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Windows Logon Application"
EndTask "explorer.exe"

UniMsgBox ChrW$(&H110) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H44) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H58) & ChrW$(&H6F) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H21), vbOKOnly, "Thông Báo", frmMain.hWnd
End Sub

Sub KillAmg()
On Error Resume Next
EndTask "amg.exe"
SetAttr "C:\Program Files\AntiMalwareGuard\amg.exe", vbNormal
DeleteFile "C:\Program Files\AntiMalwareGuard\amg.exe"
DeleteFile "C:\Program Files\AntiMalwareGuard\WL.dat"
DeleteFile "C:\Program Files\AntiMalwareGuard\BL.dat"
DeleteFile "C:\Documents and Settings\All Users\Start Menu\Programs\AntiMalwareGuard\AntiMalwareGuard.lnk"

DeleteKey HKEY_CURRENT_USER, "Software\{5222008A-DD62-49c7-A735-7BD18ECC7350}"
DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "AntiMalwareGuard"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Connections", "SavedLegacySettings"
DeleteKey HKEY_CURRENT_USER, "Software\AntiMalwareGuard"

EndTask "explorer.exe"
UniMsgBox ChrW$(&H110) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H44) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H58) & ChrW$(&H6F) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H21), vbOKOnly, "Thông Báo", frmMain.hWnd
End Sub

Sub KillIexplorer()
On Error Resume Next
EndTask "explorer.exe"
EndTask "IEXPLORER.EXE"
EndTask "WORD.EXE"

SetAttr "C:\WINDOWS\system32\IEXPLORER.exe", vbNormal
SetAttr "C:\WINDOWS\IEXPLORER.exe", vbNormal
SetAttr "C:\WINDOWS\system32\WORD.exe", vbNormal
SetAttr "C:\WINDOWS\system32\autorun.ini", vbNormal

DeleteFile "C:\WINDOWS\system32\IEXPLORER.exe"
DeleteFile "C:\WINDOWS\IEXPLORER.exe"
DeleteFile "C:\WINDOWS\system32\WORD.exe"
DeleteFile "C:\WINDOWS\system32\autorun.ini"

DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Connections", "SavedLegacySettings"
SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Userinit", "C:\WINDOWS\system32\userinit.exe"
SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", "explorer.exe"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NofolderOptions"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools"
DeleteValue HKEY_LOCAL_MACHINE, "SYSTEM\ControlSet001\Services\Schedule", "AtTaskMaxHours"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "Yahoo Messengger"
EndTask "explorer.exe"
UniMsgBox ChrW$(&H110) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H44) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H58) & ChrW$(&H6F) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H21), vbOKOnly, "Thông Báo", frmMain.hWnd

End Sub
Sub KillPCPC()
On Error Resume Next
EndTask "pcpc.exe"

SetAttr "C:\Program Files\PCPrivacyCleaner\pcpc.exe", vbNormal

DeleteFile "C:\Program Files\PCPrivacyCleaner\pcpc.exe"
DeleteFile "C:\Documents and Settings\All Users\Start Menu\Programs\PCPrivacyCleaner\PCPrivacyCleaner.lnk"
DeleteFile "C:\Documents and Settings\All Users\Start Menu\Programs\PCPrivacyCleaner\Uninstall PCPrivacyCleaner.lnk"
DeleteKey HKEY_CURRENT_USER, "Software\PCPrivacyCleaner"
DeleteKey HKEY_CURRENT_USER, "Software\{65DE966D-11D1-4bb1-BF7E-B8A273514DAF}"
DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "PCPrivacyCleaner"
DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\PCPrivacyCleaner"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Connections", "SavedLegacySettings"
EndTask "explorer.exe"
UniMsgBox ChrW$(&H110) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H44) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H58) & ChrW$(&H6F) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H21), vbOKOnly, "Thông Báo", frmMain.hWnd

End Sub
Sub KillScvhosti()
On Error Resume Next
EndTask "scvhosti.exe"
SetAttr "C:\WINDOWS\System32\scvhosti.exe", vbNormal
SetAttr "C:\WINDOWS\scvhosti.exe", vbNormal
SetAttr "C:\WINDOWS\system32\anhui.exe", vbNormal
SetAttr "C:\WINDOWS\system32\autorun.ini", vbNormal

DeleteFile "C:\WINDOWS\system32\scvhosti.exe"
DeleteFile "C:\WINDOWS\scvhosti.exe"
DeleteFile "C:\WINDOWS\system32\anhui.exe"
DeleteFile "C:\WINDOWS\system32\autorun.ini"



SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Userinit", "C:\WINDOWS\system32\userinit.exe"
SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", "explorer.exe"

DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "Yahoo Messengger"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NofolderOptions"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools"

DeleteValue HKEY_LOCAL_MACHINE, "SYSTEM\ControlSet001\Services\Schedule", "AtTaskMaxHours"

DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Connections", "SavedLegacySettings"
EndTask "explorer.exe"
UniMsgBox ChrW$(&H110) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H44) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H58) & ChrW$(&H6F) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H21), vbOKOnly, "Thông Báo", frmMain.hWnd

End Sub


Sub KillTaskmsg()
On Error Resume Next
EndTask "taskmsg.exe"
SetAttr "C:\WINDOWS\taskmsg.exe", vbNormal
SetAttr "C:\WINDOWS\save.txt", vbNormal

DeleteFile "C:\WINDOWS\taskmsg.exe"
DeleteFile "C:\WINDOWS\save.txt"

DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\ShellNoRoam\MUICache", "C:\WINDOWS\taskmsg.exe"
DeleteValue HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "WinTask"
DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "WinTask"

UniMsgBox ChrW$(&H110) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H44) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H58) & ChrW$(&H6F) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H21), vbOKOnly, "Thông Báo", frmMain.hWnd

EndTask "explorer.exe"
End Sub
Sub KillEVShuttle()
On Error Resume Next
Set colItems = GetObject("winmgmts:\root\CIMV2").ExecQuery("SELECT * FROM Win32_Process")
   For Each objitem In colItems
      Dim cap
      Dim exe
      Dim id
      cap = objitem.Caption
      exe = objitem.ExecutablePath
      id = objitem.ProcessId
      If cap = "svchost.exe" And exe = "C:\WINDOWS\svchost.exe" Then Shell "taskkill /f /im " & id
   Next
EndTask "WINDOWS.exe"
SetAttr "C:\WINDOWS\svchost.exe", vbNormal
SetAttr "C:\Folder.exe", vbNormal
SetAttr "C:\autorun.inf", vbNormal
SetAttr "D:\Folder.exe", vbNormal
SetAttr "D:\autorun.inf", vbNormal
SetAttr "E:\Folder.exe", vbNormal
SetAttr "E:\autorun.inf", vbNormal
SetAttr "F:\Folder.exe", vbNormal
SetAttr "F:\autorun.inf", vbNormal


DeleteFile "C:\folder.exe"
DeleteFile "C:\autorun.inf"
DeleteFile "D:\folder.exe"
DeleteFile "D:\autorun.inf"
DeleteFile "E:\folder.exe"
DeleteFile "E:\autorun.inf"
DeleteFile "F:\folder.exe"
DeleteFile "F:\autorun.inf"
DeleteFile "C:\WINDOWS\svchost.exe"
EndTask "explorer.exe"
DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "ctfm.exe"
UniMsgBox ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1B0) & ChrW$(&H1EE3) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H78) & ChrW$(&H6F) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H2C) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H75) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&H68) & ChrW$(&H69) & ChrW$(&HEA) _
& ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&HE2) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H6C) & ChrW$(&HE0) & ChrW$(&H20) & ChrW$(&H31) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H6C) & ChrW$(&HE2) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H46) & ChrW$(&H69) & ChrW$(&H6C) & ChrW$(&H65) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EB1) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) _
 & ChrW$(&H63) & ChrW$(&HE1) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H1EA1) & ChrW$(&H6F) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&HE1) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H46) & ChrW$(&H69) & ChrW$(&H6C) & ChrW$(&H65) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&H69) & ChrW$(&H72) & ChrW$(&H75) & ChrW$(&H73) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&HF3) & ChrW$(&H20) & ChrW$(&H49) & ChrW$(&H63) & ChrW$(&H6F) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H67) & ChrW$(&H69) & ChrW$(&H1ED1) _
 & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H1EE7) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H46) & ChrW$(&H6F) & ChrW$(&H6C) & ChrW$(&H64) & ChrW$(&H65) & ChrW$(&H72) & ChrW$(&H2E) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&HEC) & ChrW$(&H20) & ChrW$(&H76) & ChrW$(&H1EAD) & ChrW$(&H79) & ChrW$(&H20) & ChrW$(&H111) & ChrW$(&H1EC3) & ChrW$(&H20) & ChrW$(&H78) & ChrW$(&HF3) & ChrW$(&H61) & ChrW$(&H20) _
& ChrW$(&H68) & ChrW$(&H1EBF) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H1EA5) & ChrW$(&H75) & ChrW$(&H20) & ChrW$(&H76) & ChrW$(&H1EBF) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H62) & ChrW$(&H1EA1) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H1EA7) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H78) & ChrW$(&HF3) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H1EA5) _
& ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H1EA3) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&HE1) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H46) & ChrW$(&H69) & ChrW$(&H6C) & ChrW$(&H65) & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&HF3) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&HEA) & ChrW$(&H6E) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H72) & ChrW$(&HF9) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H76) & ChrW$(&H1EDB) & ChrW$(&H69) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H68) & ChrW$(&H1B0) & ChrW$(&H20) & ChrW$(&H6D) & ChrW$(&H1EE5) & ChrW$(&H63) _
 & ChrW$(&H20) & ChrW$(&H63) & ChrW$(&H68) & ChrW$(&H1EE9) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H6E) & ChrW$(&HF3) & ChrW$(&H2E) & ChrW$(&H20) & ChrW$(&H56) & ChrW$(&HED) & ChrW$(&H20) & ChrW$(&H64) & ChrW$(&H1EE5) & ChrW$(&H3A) & ChrW$(&H20) & ChrW$(&H58) & ChrW$(&HF3) & ChrW$(&H61) & ChrW$(&H20) & ChrW$(&H66) & ChrW$(&H69) & ChrW$(&H6C) & ChrW$(&H65) & ChrW$(&H20) _
 & ChrW$(&H57) & ChrW$(&H49) & ChrW$(&H4E) & ChrW$(&H44) & ChrW$(&H4F) & ChrW$(&H57) & ChrW$(&H53) & ChrW$(&H2E) & ChrW$(&H65) & ChrW$(&H78) & ChrW$(&H65) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H72) & ChrW$(&H6F) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H20) & ChrW$(&H74) & ChrW$(&H68) & ChrW$(&H1B0) & ChrW$(&H20) & ChrW$(&H6D) & ChrW$(&H1EE5) & ChrW$(&H63) & ChrW$(&H20) & ChrW$(&H57) & ChrW$(&H49) & ChrW$(&H4E) & ChrW$(&H44) & ChrW$(&H4F) & ChrW$(&H57) & ChrW$(&H53) & ChrW$(&H2E), vbOKOnly, "Thông Báo", frmMain.hWnd
End Sub

Sub KillTaquito()
On Error Resume Next

SetAttr "C:\autorun.inf", vbNormal
SetAttr "D:\autorun.inf", vbNormal
SetAttr "E:\autorun.inf", vbNormal
SetAttr "F:\autorun.inf", vbNormal

SetAttr "C:\RESTORE\S-1-5-21-1482476501-1644491937-682003330-1013\Desktop.ini", vbNormal
SetAttr "D:\RESTORE\S-1-5-21-1482476501-1644491937-682003330-1013\Desktop.ini", vbNormal
SetAttr "E:\RESTORE\S-1-5-21-1482476501-1644491937-682003330-1013\Desktop.ini", vbNormal
SetAttr "F:\RESTORE\S-1-5-21-1482476501-1644491937-682003330-1013\Desktop.ini", vbNormal

SetAttr "C:\RESTORE\S-1-5-21-1482476501-1644491937-682003330-1013\Taquito.exe", vbNormal
SetAttr "D:\RESTORE\S-1-5-21-1482476501-1644491937-682003330-1013\Taquito.exe", vbNormal
SetAttr "E:\RESTORE\S-1-5-21-1482476501-1644491937-682003330-1013\Taquito.exe", vbNormal
SetAttr "F:\RESTORE\S-1-5-21-1482476501-1644491937-682003330-1013\Taquito.exe", vbNormal

DeleteFile "C:\autorun.inf"
DeleteFile "D:\autorun.inf"
DeleteFile "E:\autorun.inf"
DeleteFile "F:\autorun.inf"

DeleteFile "C:\RESTORE\S-1-5-21-1482476501-1644491937-682003330-1013\Desktop.ini"
DeleteFile "D:\RESTORE\S-1-5-21-1482476501-1644491937-682003330-1013\Desktop.ini"
DeleteFile "E:\RESTORE\S-1-5-21-1482476501-1644491937-682003330-1013\Desktop.ini"
DeleteFile "F:\RESTORE\S-1-5-21-1482476501-1644491937-682003330-1013\Desktop.ini"

DeleteFile "C:\RESTORE\S-1-5-21-1482476501-1644491937-682003330-1013\Taquito.exe"
DeleteFile "D:\RESTORE\S-1-5-21-1482476501-1644491937-682003330-1013\Taquito.exe"
DeleteFile "E:\RESTORE\S-1-5-21-1482476501-1644491937-682003330-1013\Taquito.exe"
DeleteFile "F:\RESTORE\S-1-5-21-1482476501-1644491937-682003330-1013\Taquito.exe"
EndTask "explorer.exe"
UniMsgBox ChrW$(&H110) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H44) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H58) & ChrW$(&H6F) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H21), vbOKOnly, "Thông Báo", frmMain.hWnd

End Sub

Sub KillTiepTuc()
On Error Resume Next

EndTask "explorer.exe"
EndTask "HayTiepTuc.exe"
Set colItems = GetObject("winmgmts:\root\CIMV2").ExecQuery("SELECT * FROM Win32_Process")
   For Each objitem In colItems
      Dim cap
      Dim exe
      Dim id
      cap = objitem.Caption
      exe = objitem.ExecutablePath
      id = objitem.ProcessId
      If cap = "HayTiepTuc.exe" And exe = "C:\WINDOWS\system32\Sys\HayTiepTuc.exe" Then Shell "taskkill /f /im " & id
   Next
Shell "taskkill /f /im haytieptuc.exe"
SetAttr "C:\WINDOWS\System32\Sys\HayTiepTuc.exe", vbNormal
SetAttr "C:\WINDOWS\System32\Sys\HayTiepTuc.001", vbNormal
SetAttr "C:\WINDOWS\System32\Sys\HayTiepTuc.002", vbNormal
SetAttr "C:\WINDOWS\System32\Sys\HayTiepTuc.003", vbNormal
SetAttr "C:\WINDOWS\System32\Sys\HayTiepTuc.004", vbNormal
SetAttr "C:\WINDOWS\System32\Sys\HayTiepTuc.005", vbNormal
SetAttr "C:\WINDOWS\System32\Sys\HayTiepTuc.006", vbNormal
SetAttr "C:\WINDOWS\System32\Sys\HayTiepTuc.007", vbNormal


DeleteFile "C:\WINDOWS\System32\Sys\HayTiepTuc.exe"
DeleteFile "C:\WINDOWS\System32\Sys\HayTiepTuc.001"
DeleteFile "C:\WINDOWS\System32\Sys\HayTiepTuc.002"
DeleteFile "C:\WINDOWS\System32\Sys\HayTiepTuc.003"
DeleteFile "C:\WINDOWS\System32\Sys\HayTiepTuc.004"
DeleteFile "C:\WINDOWS\System32\Sys\HayTiepTuc.005"
DeleteFile "C:\WINDOWS\System32\Sys\HayTiepTuc.006"
DeleteFile "C:\WINDOWS\System32\Sys\HayTiepTuc.007"

DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\ShellNoRoam\MUICache", "C:\WINDOWS\system32\Sys\HayTiepTuc.exe"

EndTask "explorer.exe"
UniMsgBox ChrW$(&H110) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H44) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H58) & ChrW$(&H6F) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H21), vbOKOnly, "Thông Báo", frmMain.hWnd

End Sub
Sub KillMegabyte()
On Error Resume Next
Shell "taskkill /f /im explorer.exe"
Shell "taskkill /f /im 250mb.exe"
Shell "taskkill /f /im megabyte.exe"
EndTask "250mb.exe"
EndTask "megabyte.exe"

SetAttr "C:\WINDOWS\250mb.exe", vbNormal
SetAttr "C:\WINDOWS\megabyte.exe", vbNormal

DeleteFile "C:\WINDOWS\250mb.exe"
DeleteFile "C:\WINDOWS\megabyte.exe"

SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Userinit", "C:\WINDOWS\system32\userinit.exe"
SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", "explorer.exe"
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", 0
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr", 0
SaveDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions", 0
DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Folder\Hidden\ShowAll", "CheckedValue"
SaveDWORD HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Folder\Hidden\ShowAll", "CheckedValue", 1

SaveString HKEY_CLASSES_ROOT, "exefile\shell\open\command", "", ChrW$(&H22) & ChrW$(&H25) & ChrW$(&H31) & ChrW$(&H22) & ChrW$(&H20) & ChrW$(&H25) & ChrW$(&H2A)
EndTask "explorer.exe"
Shell "explorer.exe"
UniMsgBox ChrW$(&H110) & ChrW$(&HE3) & ChrW$(&H20) & ChrW$(&H44) & ChrW$(&H69) & ChrW$(&H1EC7) & ChrW$(&H74) & ChrW$(&H20) & ChrW$(&H58) & ChrW$(&H6F) & ChrW$(&H6E) & ChrW$(&H67) & ChrW$(&H21), vbOKOnly, "Thông Báo", frmMain.hWnd

End Sub

'SCVVHSOT
Sub KillSxS()
On Error Resume Next
EndTask "sxs.exe"
EndTask "explorer.exe"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\Run", "AudioMan"
DeleteKey HKEY_LOCAL_MACHINE, "Classes\.sm1"
DeleteKey HKEY_LOCAL_MACHINE, "\SOFTWARE\Classes\sm1_Auto_File\shell\open\command"
DeleteKey HKEY_LOCAL_MACHINE, "\SOFTWARE\Classes\sm1_Auto_File\shell\open"
DeleteKey HKEY_LOCAL_MACHINE, "\SOFTWARE\Classes\sm1_Auto_File\shell"
DeleteKey HKEY_LOCAL_MACHINE, "\SOFTWARE\Classes\sm1_Auto_File"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\ShellNoRoam\MUICache", "C:\WINDOWS\system32\sc.exe"

SetAttr "C:\WINDOWS\system32\Explorer.sm1", vbNormal
DeleteFile "C:\WINDOWS\system32\Explorer.sm1"

UniMsgBox ChrW(272) & ChrW(227) & " di" & ChrW(7879) & "t xong."

End Sub
Function GetUser()
    GetUser = Environ$("username")
End Function
Sub KillSCVVHSOT()
On Error Resume Next
Shell "taskkill /f /im SCVVHSOT.exe"
Shell "taskkill /f /im explorer.exe"

EndTask "SCVVHSOT.exe"
EndTask "explorer.exe"


DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\policies\system", "EnableLUA"
DeleteKey HKEY_LOCAL_MACHINE, "SYSTEM\ControlSet001\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile\AuthorizedApplications\List"

DeleteKey HKEY_CURRENT_USER, "Software\Administrator914\-993627007"
DeleteKey HKEY_CURRENT_USER, "Software\Administrator914"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "Yahoo Messengger"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\WorkgroupCrawler\Shares", "shared"
CleanReg

DeleteFile "C:\Documents and Settings\" & GetUser & "\Local Settings\Temp\0022873B_Rar\SCVVHSOT.exe"
Kill "C:\Documents and Settings\" & GetUser & "\Local Settings\Temp\*.exe"

SetAttr "C:\WINDOWS\system32\SCVVHSOT.exe", vbNormal
SetAttr "C:\WINDOWS\SCVVHSOT.exe", vbNormal
SetAttr "C:\WINDOWS\system32\blastclnnn.exe", vbNormal
SetAttr "C:\WINDOWS\system32\autorun.ini", vbNormal
SetAttr "C:\WINDOWS\temp\Perflib_Perfdata_21c.dat", vbNormal
SetAttr "C:\WINDOWS\System32\Drivers\hlknon.sys", vbNormal


DeleteFile "C:\WINDOWS\system32\SCVVHSOT.exe"
DeleteFile "C:\WINDOWS\SCVVHSOT.exe"
DeleteFile "C:\WINDOWS\system32\blastclnnn.exe"
DeleteFile "C:\WINDOWS\system32\autorun.ini"
DeleteFile "C:\WINDOWS\temp\Perflib_Perfdata_21c.dat"
DeleteFile "C:\WINDOWS\System32\Drivers\hlknon.sys"

Kill "C:\*.pif"
Shell "attrib -s -h -r C:\*.exe"
Kill "D:\*.pif"
Shell "attrib -s -h -r D:\*.exe"
Kill "F:\*.pif"
Shell "attrib -s -h -r F:\*.exe"
Kill "E:\*.pif"
Shell "attrib -s -h -r E:\*.exe"

Shell "explorer.exe"


UniMsgBox ChrW(272) & ChrW(227) & " di" & ChrW(7879) & "t xong, h" & ChrW(227) & "y Log Off l" & ChrW(7841) & "i m" & ChrW(225) & "y " & ChrW(273) & ChrW(7875) & " ho" & ChrW(224) & "n t" & ChrW(7845) & "t c" & ChrW(244) & "ng vi" & ChrW(7879) & "c x" & ChrW(243) & "a Virus ra kh" & ChrW(7887) & "i m" & ChrW(225) & "y t" & ChrW(237) & "nh.", vbOKOnly, "Thông Báo", frmMain.hWnd

End Sub


Sub KillWin1ogon()
On Error Resume Next
EndTask "win1ogon.exe"
DeleteKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\Notify\win1ogon"
DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "win1ogon"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\ShellNoRoam\MUICache", "C:\WINDOWS\system32\win1ogon.exe"

SetAttr "C:\WINDOWS\system32\win1ogon.exe", vbNormal
SetAttr "C:\WINDOWS\system32\win1ogon.dll", vbNormal
DeleteFile "C:\WINDOWS\system32\win1ogon.exe"
DeleteFile "C:\WINDOWS\system32\win1ogon.dll"

EndTask "explorer.exe"
UniMsgBox ChrW(272) & ChrW(227) & " di" & ChrW(7879) & "t xong. Ch" & ChrW(250) & " " & ChrW(221) & ": !!! " & ChrW(272) & ChrW(226) & "y l" & ChrW(224) & " m" & ChrW(7897) & "t lo" & ChrW(7841) & "i Virus l" & ChrW(226) & "y file .exe n" & ChrW(234) & "n " & ChrW(273) & ChrW(7875) & " " & ChrW(273) & ChrW(7843) & "m b" & ChrW(7843) & "o m" & ChrW(225) & "y t" & ChrW(237) & "nh kh" & ChrW(244) & "ng b" & ChrW(7883) & " nhi" & ChrW(7877) & "m l" & ChrW(7841) & "i l" & ChrW(7847) & "n n" & ChrW(7919) & "a, b" & ChrW(7841) & "n h" & ChrW(227) & "y th" & ChrW(7921) & "c hi" & ChrW(7879) & "n sao l" & ChrW(432) & "u v" & ChrW(224) & " Repair Windows " & ChrW(273) & ChrW(7875) & " di" & ChrW(7879) & "t t" & ChrW(7853) & "n g" & ChrW(7889) & "c Virus n" & ChrW(224) & "y..", vbOKOnly, "Thông Báo", frmMain.hWnd

End Sub







Sub Kill2FIY()
On Error Resume Next
EndTask "explorer.exe"
DeleteValue HKEY_LOCAL_MACHINE, "SYSTEM\ControlSet001\Services\KAVsys", "ImagePath"
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "cdoosoft"

SetAttr "C:\WINDOWS\system32\drivers\klif.sys", vbNormal
SetAttr "C:\WINDOWS\system32\olhrwef.exe", vbNormal
SetAttr "C:\WINDOWS\system32\nmdfgds0.dll", vbNormal

DeleteFile "C:\WINDOWS\system32\drivers\klif.sys"
DeleteFile "C:\WINDOWS\system32\olhrwef.exe"
DeleteFile "C:\WINDOWS\system32\nmdfgds0.dll"

SetAttr "C:\autorun.inf", vbNormal
SetAttr "C:\2fiy.bat", vbNormal
SetAttr "D:\autorun.inf", vbNormal
SetAttr "D:\2fiy.bat", vbNormal
SetAttr "E:\autorun.inf", vbNormal
SetAttr "E:\2fiy.bat", vbNormal
SetAttr "F:\autorun.inf", vbNormal
SetAttr "F:\2fiy.bat", vbNormal

xcdqj.exe


DeleteFile "C:\autorun.inf"
DeleteFile "C:\2fiy.bat"
DeleteFile "C:\xcdqj.exe"

DeleteFile "D:\autorun.inf"
DeleteFile "D:\2fiy.bat"
DeleteFile "C:\xcdqj.exe"

DeleteFile "E:\autorun.inf"
DeleteFile "E:\2fiy.bat"
DeleteFile "C:\xcdqj.exe"

DeleteFile "F:\autorun.inf"
DeleteFile "F:\2fiy.bat"
DeleteFile "C:\xcdqj.exe"

Kill "C:\*.pif"
Kill "D:\*.pif"
Kill "E:\*.pif"
Kill "F:\*.pif"


Kill "D:\*.cmd"
Kill "E:\*.cmd"
Kill "F:\*.cmd"

Shell "attrib -s -h -r C:\*.exe"
Shell "attrib -s -h -r D:\*.exe"
Shell "attrib -s -h -r E:\*.exe"

Shell "explorer.exe"
EndTask "cmd.exe"
UniMsgBox ChrW(272) & ChrW(227) & " di" & ChrW(7879) & "t xong. Tuy nhi" & ChrW(234) & "n m" & ChrW(7897) & "t s" & ChrW(7889) & " ch" & ChrW(7913) & "c n" & ChrW(259) & "ng c" & ChrW(7911) & "a Windows c" & ChrW(243) & " th" & ChrW(7875) & " v" & ChrW(7851) & "n b" & ChrW(7883) & " kh" & ChrW(243) & "a do c" & ChrW(225) & "c Drivers c" & ChrW(7911) & "a Virus c" & ChrW(242) & "n " & ChrW(273) & ChrW(7875) & " l" & ChrW(7841) & "i.", vbOKOnly, "Thông Báo", frmMain.hWnd
End Sub
