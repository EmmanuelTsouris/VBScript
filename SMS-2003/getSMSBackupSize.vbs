
'
' Get the SMS Backup Size
' 7/2/2009
'

Wscript.echo "folderPath,folderName,Size,fileCount,folderCount

ShowFolderDetails "\\path\to\SMSBackup"

sub ShowFolderDetails(folderPath)
	dim oFS, oFolder

	set oFS = WScript.CreateObject("Scripting.FileSystemObject")
	set oF = oFS.GetFolder(folderPath)

	Wscript.echo folderPath & "," & oF.Name & "," & oF.Size, oF.Files.Count, oF.Subfolders.count

end sub