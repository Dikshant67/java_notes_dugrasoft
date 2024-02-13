'1. Create Folder at location C:\ with your name
option Explicit
dim fileobj
dim src
set fileobj = CreateObject("Scripting.FileSystemObject")
src="C:\Dikshant"
if fileobj.folderexists(src) then
  msgbox "Folder Already Exists"
 Else
 fileobj.CreateFolder src
 msgbox "folder created Successfully at C:\"
end if 
set fileobj=Nothing