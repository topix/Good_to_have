Dim FSO, objShell, wshShell
Set FSO = CreateObject("Scripting.FileSystemObject")
Set objShell = Wscript.CreateObject("WScript.Shell")
Set wshShell = CreateObject( "WScript.Shell" )

const destFolderProcess = "C:\Utils\script_Process_restart\"
const fileScript = "reset.ps1"
const fileRunVBS = "run.vbs"
const fileIcon = "restart.ico"

If NOT (FSO.FolderExists("C:\Utils\")) Then
	FSO.CreateFolder "C:\Utils\"
End If
If NOT (FSO.FolderExists("C:\Utils\script_Process_restart\")) Then
	FSO.CreateFolder "C:\Utils\script_Process_restart\"
End If

'Copy all files that belong in "C:\Utils\script_Process_restart\"
FSO.CopyFile fileScript, destFolderProcess + fileScript, True
FSO.CopyFile fileRunVBS, destFolderProcess + fileRunVBS, True
FSO.CopyFile fileIcon, destFolderProcess + fileIcon, True
On Error Resume Next

Set lnk = objShell.CreateShortcut("C:\Utils\script_Process_restart\Process-Restart.LNK")
lnk.TargetPath = "C:\Utils\script_Process_restart\run.vbs"
lnk.Arguments = ""
lnk.Description = "Process-Restart"
lnk.WorkingDirectory = "C:\Utils\script_Process_restart"
lnk.IconLocation = "C:\Utils\script_Process_restart\restart.ico"
lnk.Save
'Clean up 
Set lnk = Nothing

'Which version are you running?
Set shell = CreateObject("WScript.Shell")
Set getOSVersion = shell.exec("%comspec% /c ver")
version = getOSVersion.stdout.readall
Select Case True
    Case InStr(version, "n 5.") > 1 : GetOS = "XP"
    Case InStr(version, "n 6.1.") > 1 : GetOS = "W7"
    Case InStr(version, "n 10.") > 1 : GetOS = "Windows 10"
    Case Else : GetOS = "Unknown"
End Select 

'Ok, then i put it in the right directory.
strUserName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )
If FSO.FileExists ("C:\Utils\script_Process_restart\Process-Restart.LNK") then
    If GetOS = "XP" then
        If FSO.FileExists ("C:\Documents and Settings\" & strUserName & "\Desktop\Process-Restart.LNK") then
		  FSO.DeleteFile "C:\Documents and Settings\" & strUserName & "\Desktop\Process-Restart.LNK"
        End If
    ElseIf GetOS = "W7" then
        If FSO.FileExists ("C:\Users\Emil\" & strUserName & "\Desktop\Process-Restart.LNK") then
		FSO.DeleteFile "C:\Users\Emil\" & strUserName & "\Desktop\Process-Restart.LNK"
	   End If
    End If
    If GetOS = "XP" then
        FSO.CopyFile "C:\Utils\script_Process_restart\Process-Restart.LNK", "C:\Documents and Settings\" & strUserName & "\Desktop\Process-Restart.LNK"
    ElseIf GetOS = "W7" then
        FSO.CopyFile "C:\Utils\script_Process_restart\Process-Restart.LNK", "C:\Users\Emil\" & strUserName & "\Desktop\Process-Restart.LNK"
    End If
End If




Set FSO = Nothing
Set objShell = Nothing


