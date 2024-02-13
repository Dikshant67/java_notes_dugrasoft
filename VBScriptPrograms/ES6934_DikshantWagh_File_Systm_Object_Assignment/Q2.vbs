'2. Create .txt file at location C:\Vrushali\Testing with name "MyFirstTextFile.txt“
option Explicit
dim fileobj
set fileobj= CreateObject("Scripting.FileSystemObject")


if fileobj.folderexists("C:\Vrushali") then 
 msgbox "Folder Already Exists"
Else
 fileobj.createfolder "C:\Vrushali"
 msgbox "Folder Created C:\Vrushali"
end if

if fileobj.folderexists("C:\Vrushali\Testing") then 
 msgbox "Folder already Exists"
Else
 fileobj.createfolder "C:\Vrushali\Testing"
  msgbox "Folder Created C:\Vrushali\Testing"
end if


if fileobj.fileexists("C:\Vrushali\Testing\MyFirstTextFile.txt") then 
 msgbox "MyFirstTextFile.txt already Exists"
Else
 fileobj.CreateTextFile "C:\Vrushali\Testing\MyFirstTextFile.txt" 
 msgbox "MyFirstTextFile.txt created Succussfully at Location C:\Vrushali\Testing\"
end if


set fileobj=Nothing
