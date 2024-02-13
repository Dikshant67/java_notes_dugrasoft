'4. Copy the text from the same file ("MyFirstTextFile.txt") and create an another text file with name "Result.txt" where you can paste the copied data
option Explicit
dim fileobj,fileobj1,fileobj2
dim str
set fileobj= CreateObject("Scripting.FileSystemObject")
const forReading=1
const forWriting=2
set fileobj1= fileobj.openTextFile("C:\Vrushali\Testing\MyFirstTextFile.txt",forReading)
str=fileobj1.ReadAll
fileobj.CreateTextFile "C:\Vrushali\Testing\Result.txt"
set fileobj2=fileobj.openTextFile("C:\Vrushali\Testing\Result.txt",forWriting)
fileobj2.WriteLine(str)
fileobj1.close
fileobj2.close

set fileobj=Nothing