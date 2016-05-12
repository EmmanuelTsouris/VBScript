
'
' Opens a command window, sends the keys "exit" and then the window closes
' 
'

Option Explicit
Dim objShell, Racey, intCount
Set objShell = CreateObject("WScript.Shell")
objShell.Run "cmd" 
Wscript.Sleep 1000
objShell.SendKeys "exit~"
WScript.Quit 
' End of Example SendKeys VBScript
