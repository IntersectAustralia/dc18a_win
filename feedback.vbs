'--- START ---
Option Explicit
On Error Resume Next
Dim WSHShell, msg, MyBox, title, strCmd, objNet
Set WSHShell = WScript.CreateObject("WScript.Shell")
Set objNet = WScript.CreateObject("WScript.NetWork")
strCmd = "%ComSpec% /c """ & "%ProgramFiles(x86)%\Internet Explorer\iexplore.exe" & """ http://dc18a-staging.intersect.org.au/experiment_feedbacks/new?login_id=" & objNet.UserName
WSHShell.Run strCmd, 0, true

'Output message for dialog box

'msg = "Experiment successfully finished?" & vbCR
'title = "Microbial Imaging Facility"
'MyBox = MsgBox (msg, 4132, title)

Set WshShell = Nothing
'--- END ---
