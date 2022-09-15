
Dim WshShell, strCurDir
Set WshShell = CreateObject("WScript.Shell")
strCurDir = WshShell.CurrentDirectory

Set ObjFso = CreateObject("\scripting.FileSystemObject")

Set Folder = objFSO.GetFolder(strCurDir)

version = InputBox("Enter Revision Number :")


function cfolder(fname, objFso)
if objFSO.FolderExists(fname)  Then
	WScript.Echo "Folder exists "&fname

else
	objFSO.CreateFolder(fname)
	WScript.Echo "Created "&fname
end if
end function

'4506


Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile("plantlist.txt",1)
Dim strLine
do while not objFileToRead.AtEndOfStream
     fname = objFileToRead.ReadLine()
     c = cfolder(fname, objFso)
'     WScript.Echo fname
     'Do something with the line
loop
objFileToRead.Close
Set objFileToRead = Nothing





