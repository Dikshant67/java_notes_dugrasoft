'5. Now move newly created "Result" file at location C:\Vrushali\Testing\Result
option Explicit
dim fileobj,fileobj1
dim str
dim sourcePath,destPath
set fileobj= CreateObject("Scripting.FileSystemObject")
sourcePath="C:\Vrushali\Testing\Result.txt"
destPath="C:\Vrushali\Testing\Result\"

fileobj.createFolder("C:\Vrushali\Testing\Result")
 if fileobj.FileExists(sourcePath) Then
    fileobj.MoveFile sourcePath,destPath
 Else
	msgbox "File Not Exist"
 end if 
set fileobj=Nothing