dim oFS, oFolder

set oFS = WScript.CreateObject("Scripting.FileSystemObject")
set oFolder = oFS.GetFolder("c:\temp")

ShowFolderDetails oFolder

sub ShowFolderDetails(oF)

	dim F

	Wscript.echo oF.Name & "," & oF.Size, oF.Files.Count, oF.Subfolders.count, oF.Size

	'for each F in oF.Subfolders
       ' 	ShowFolderDetails(F)
    '	next

end sub