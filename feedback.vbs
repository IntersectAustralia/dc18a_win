'--- START ---
Option Explicit
On Error Resume Next
Dim WSHShell, msg, MyBox, title, strCmd
Set WSHShell = WScript.CreateObject("WScript.Shell")
strCmd = "%ComSpec% /c """ & "%ProgramFiles(x86)%\Internet Explorer\iexplore.exe" & """ http://dc18a-staging.intersect.org.au"
WSHShell.Run strCmd, 0, true

'Output message for dialog box

'msg = "Experiment successfully finished?" & vbCR
'title = "Microbial Imaging Facility"
'MyBox = MsgBox (msg, 4132, title)

Set WshShell = Nothing
'--- END ---