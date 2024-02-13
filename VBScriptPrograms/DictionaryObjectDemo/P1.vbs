Option Explicit
Dim obj_datadict 'created a variable 
Dim obj_item
dim obj_key
dim i 	
set obj_datadict=CreateObject("Scripting.Dictionary")
obj_datadict.add "b" , "Mango"
obj_datadict.add "a" , "Apple"
obj_datadict.add "c" , "Guava"

' if obj_datadict.Exists("c") then 
 ' msgbox "Specified Key Exists"
 ' Else 
  ' msgbox "Specified Key doesnot Exists"
' end if 

obj_datadict.remove("c")
obj_item= obj_datadict.items
obj_key=obj_datadict.keys
for i=0 to obj_datadict.Count-1

  msgbox(obj_item(i)&"   "&obj_key(i))
 Next
 