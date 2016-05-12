
'
' Gets the full name of the vbscript
'

Set fso = CreateObject("Scripting.FileSystemObject") 

wscript.Echo wscript.ScriptFullName 

wscript.Echo fso.GetParentFolderName(wscript.ScriptFullName) 

