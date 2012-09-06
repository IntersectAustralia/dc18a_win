Set oShell = CreateObject ("WScript.Shell")
Set objNet = CreateObject ("WScript.NetWork")
 
Dim strCmd
strCmd = "%ComSpec% /c """ & "%ProgramFiles(x86)%\Internet Explorer\iexplore.exe" & """ -k http://172.16.4.78:3000/experiments/new?login_id=" & objNet.UserName 
oShell.Run strCmd, 0, false