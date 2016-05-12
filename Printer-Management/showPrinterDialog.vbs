 
 '
 ' Shows the printer dialog box in Windows
 '
 
dim objShell
dim bReturn

set objShell = CreateObject("shell.application")
objShell.FindPrinter()

set objShell = nothing
