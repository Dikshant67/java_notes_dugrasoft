option Explicit
dim fileobj1,fileobj
dim str,str1
set fileobj1= CreateObject("Scripting.FileSystemObject")
const forReading=1
const forWriting=2
set fileobj=fileobj1.OpenTextFile("C:\MyFolder\MyDetails.txt",forReading)
str=fileobj.ReadLine
MsgBox str
Do while not fileobj.AtEndOfStream
 str1=fileobj.ReadLine
 msgbox str1
 Loop 
 set fileobj=Nothing
 set fileobj1=nothing