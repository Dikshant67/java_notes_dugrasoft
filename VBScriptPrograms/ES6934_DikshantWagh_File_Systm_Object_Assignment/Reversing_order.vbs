option Explicit
dim fileobj1,fileobj,fileobj2
dim str,str1
dim count:count=0
dim iter,i
dim arr(5)
set fileobj1= CreateObject("Scripting.FileSystemObject")
const forReading=1
const forWriting=2
set fileobj=fileobj1.OpenTextFile("C:\MyFolder\MyDetails.txt",forReading)
set fileobj2=fileobj1.OpenTextFile("C:\MyFolder\demo.txt",forWriting)

 ' str= fileobj.ReadAll
 ' msgbox str
Do until fileobj.AtEndOfStream
    arr(count)=fileobj.ReadLine
   count=count+1
 Loop 

 for i=UBound(arr) to 0 step -1
  fileobj2.WriteLine(arr(i))
 Next 
 
 set fileobj=Nothing
 set fileobj1=nothing
  set fileobj2=nothing
  MsgBox "Done"