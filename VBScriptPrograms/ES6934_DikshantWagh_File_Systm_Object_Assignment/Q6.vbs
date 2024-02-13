'6. Now delete the text file : "MyFirstTextFile.txt“ from the location.
option Explicit
dim fileobj,fileobj1
dim str
dim sourcePath,destPath
set fileobj= CreateObject("Scripting.FileSystemObject")
sourcePath="C:\Vrushali\Testing\MyFirstTextFile.txt"

 if fileobj.FileExists(sourcePath) Then
    fileobj.deletefile(sourcepath)
	msgbox "file deleted Successfully"
 Else
	msgbox "File Not Exists"
 end if 
set fileobj=Nothing