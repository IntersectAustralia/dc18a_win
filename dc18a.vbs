Set oShell = CreateObject ("WScript.Shell")
Set objNet = CreateObject ("WScript.NetWork")
 
Dim strCmd
strCmd = "%ComSpec% /c """ & "%ProgramFiles(x86)%\Internet Explorer\iexplore.exe" & """ -k http://dc18a-staging.intersect.org.au/experiments/new?login_id=" & objNet.UserName 
oShell.Run strCmd, 0, false