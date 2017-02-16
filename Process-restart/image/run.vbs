 Set oShell = CreateObject("Shell.Application")  
 oShell.ShellExecute "powershell", "-executionpolicy bypass -file C:\Utils\script_Process_restart\reset.ps1", "", "runas", 1 