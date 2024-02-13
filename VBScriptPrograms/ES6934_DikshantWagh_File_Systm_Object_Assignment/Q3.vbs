'3. Write some text under the same file - "MyFirstTextFile.txt“
option Explicit
dim fileobj,fileobj1
set fileobj= CreateObject("Scripting.FileSystemObject")
const forWriting=2
set fileobj1= fileobj.openTextFile("C:\Vrushali\Testing\MyFirstTextFile.txt",forWriting)
fileobj1.WriteLine("This text is written in MyFirstTextFile")
fileobj1.close
set fileobj=Nothing