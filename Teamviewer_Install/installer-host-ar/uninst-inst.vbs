Option Explicit
' Define variables
Dim objShell, oFSO, strUninstCD, strUninstMSI, strInstCD, strInstMSI, sScriptDir, teamViewerMSI
Dim shell, getOSVersion, version, GetOS, OsType
' Set variables
Set objShell = WScript.CreateObject( "WScript.Shell" )
Set oFSO = CreateObject("Scripting.FileSystemObject")
sScriptDir = oFSO.GetParentFolderName(WScript.ScriptFullName)
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' Get OS Version
Set shell = CreateObject("WScript.Shell")
Set getOSVersion = shell.exec("%comspec% /c ver")
version = getOSVersion.stdout.readall
Select Case True
    Case InStr(version, "n 5.") > 1 : GetOS = "XP"
    Case InStr(version, "n 6.1.") > 1 : GetOS = "W7"
    Case InStr(version, "n 10.") > 1 : GetOS = "Windows 10"
    Case Else : GetOS = "Unknown"
End Select
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
OsType = objShell.RegRead("HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\PROCESSOR_ARCHITECTURE")
'If OsType = "x86" then
'ElseIf OsType = "AMD64" then
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' Path to MSI file
Recurse oFSO.GetFolder(sScriptDir)
Sub Recurse(objFolder)
    Dim objFile, objSubFolder

    For Each objFile In objFolder.Files
        If LCase(oFSO.GetExtensionName(objFile.Name)) = "msi" Then
			teamViewerMSI = objFile.Name
        End If
    Next
End Sub
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'#########################################################################
'#########################################################################
' !!!! Question Time !!!!
' Check if OS-Host 3.1+ already installed
' IF YES Then Warn user and promt with YES or NO
'#########################################################################
'#########################################################################
' !!!! Uninstallation Time !!!!
On Error Resume Next

' Try Uninstall on all old versions ASC (With .exe)
Teamviewer
Sub Teamviewer()
	On Error Resume Next
	If OsType = "x86" Then
		'Wscript.Echo "System is x86"
		objShell.Exec "C:\Program Files\TeamViewer\Version7\uninstall.exe /S" 'Uninstall 7
		objShell.Exec "C:\Program Files\TeamViewer\Version8\uninstall.exe /S" 'Uninstall 8
		objShell.Exec "C:\Program Files\TeamViewer\Version9\uninstall.exe /S" 'Uninstall 9
		objShell.Exec "C:\Program Files\TeamViewer\uninstall.exe /S" 'Uninstall 10 +
	End If
	If OsType = "AMD64" Then
		'Wscript.Echo "System is AMD64"
		objShell.Exec "C:\Program Files (x86)\TeamViewer\Version7\uninstall.exe /S" 'Uninstall 7
		objShell.Exec "C:\Program Files (x86)\TeamViewer\Version8\uninstall.exe /S" 'Uninstall 8
		objShell.Exec "C:\Program Files (x86)\TeamViewer\Version9\uninstall.exe /S" 'Uninstall 9
		objShell.Exec "C:\Program Files (x86)\TeamViewer\uninstall.exe /S" 'Uninstall 10 +
	End If
End Sub
' Uninstall with MSI Wrapper (After .exe removed everything)
strUninstCD = "CMD /C CD /D " & sScriptDir & " & "
strUninstMSI = "msiexec /qn /uninstall " & teamViewerMSI & " /norestart"
objShell.Run strUninstCD & strUninstMSI, 1, True 'Uninstall MSI Wrapper
'Wscript.Echo.Echo strUninstCD & strUninstMSI & " & exit"

' Clean up in Windows Registry if needed
' !!!! CODE ON THE WAY !!!!

'#########################################################################
'#########################################################################
' !!!! Installation Time !!!!
' Install with MSI Wrapper (All REG keys has been purged above)
strInstCD = "CMD /C CD /D " & sScriptDir & " & "
strInstMSI = "msiexec /qn /i " & teamViewerMSI & " /norestart"
objShell.Run strInstCD & strInstMSI, 1, True 'Install with MSI Wrapper
'Wscript.Echo.Echo strInstCD & strInstMSI & " & exit"

' Prompt User and Restart
Dim objOutputFile
Set objShell = WScript.CreateObject("WScript.Shell") 
Set oFSO = CreateObject("Scripting.FileSystemObject") 
Set objOutputFile = oFSO.OpenTextFile("temp.vbs", 2, -1) 
objOutputFile.WriteLine "MsgBox ""Computer will Restart in 10sec""" 
objOutputFile.Close 
objShell.Run "temp.vbs", 1, False
WScript.Sleep 10000
oFSO.DeleteFile "temp.vbs"
'restart, waited 10 seconds, force running apps to close
objShell.Run "%comspec% /c shutdown /r /t 1 /f", , TRUE

'#########################################################################
'#########################################################################
' !!!! Call extra script below !!!!


Set objShell = Nothing