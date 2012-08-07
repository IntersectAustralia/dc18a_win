Set oShell = CreateObject ("Wscript.Shell") 
Dim strCmd
strCmd = "%ComSpec% /c """ & "%ProgramFiles(x86)%\Internet Explorer\iexplore.exe" & """ -k http://172.16.4.78:3000"
oShell.Run strCmd, 0, false