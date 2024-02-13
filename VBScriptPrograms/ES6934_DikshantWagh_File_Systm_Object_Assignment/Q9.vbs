' 9. Write a program to write your 3 line address in text file
option Explicit
dim fileobj,fileobj1
dim address
dim forWriting
const forWriting=2
address="23,New GoodWill PG"&vbNewLine&"Phase 1 Hinjawadi"&vbNewLine&"Pune"



set fileobj= CreateObject("Scripting.FileSystemObject")
if fileobj.fileexists("C:\Dikshant\MyFirstTextFile.txt") then 
 msgbox "MyAdress.txt already Exists"
Else
 fileobj.CreateTextFile "C:\Vrushali\Testing\MyFirstTextFile.txt" 
 msgbox "MyAdress.txt created Succussfully at Location C:\Vrushali\Testing\"
end if
set fileobj1= fileobj.openTextFile("C:\Vrushali\Testing\MyAdress.txt",forWriting)

do until address.atendofstream
fileobj1.WriteLine(address)
Next

fileobj1.close
set fileobj=Nothing

