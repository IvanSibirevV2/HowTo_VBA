'”рок VBScript є6:
'÷иклы For ... Next и For Each ... Next
'file_3.vbs

Dim FSO, Drive, Folders, Folder, FolderList

Set FSO = CreateObject("Scripting.FileSystemObject")  
Set Drive = FSO.GetFolder("C:\")
Set Folders = Drive.SubFolders 

For Each Folder In Folders     
	FolderList = FolderList & Folder.Name & vbCrLf  
Next 

MsgBox FolderList