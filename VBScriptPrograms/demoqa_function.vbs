Option Explicit
Dim ObjDesc,Child_Obj
Dim str_text
Dim iter

Function ValidateTextBox(Full_name,Email,Current_address,Permanant_adress)
  dim flag:flag=True
  if Browser("DEMOQA").Page("DEMOQA_3").WebEdit("Full Name").Exist(4) Then
    set ObjDesc =Description.Create
	ObjDesc("html tag").Value="p"
	Set Child_Obj=Browser("DEMOQA").Page("DEMOQA_3").ChildObjects(ObjDesc)
	For iter=0 to Child_Obj.count-1 step 1
	 str_text=Child_Obj(iter).GetROProperty("innerhtml")
	 msgbox str_text
	 msgbox instr(str_text,Full_name)
	  if not(Instr(str_text,Full_name)>0 or Instr(str_text,Email)>0 or Instr(str_text,Current_address)>0 or Instr(str_text,Permanant_adress)>0  ) then
	    flag=False
		Exit For
	  end if
    Next
  Else
    ValidateTextBox=False
  end if 
    ValidateTextBox=flag
	
end Function
Function TextFill(name,email,current_address,permanant_address)
msgbox email
Browser("DEMOQA").Page("DEMOQA").WebElement("Elements").Click
Browser("DEMOQA").Page("DEMOQA_2").WebElement("item-0").Click
Browser("DEMOQA").Page("DEMOQA_3").WebEdit("Full Name").Set name
Browser("DEMOQA").Page("DEMOQA_3").WebEdit("name@example.com").Set email
Browser("DEMOQA").Page("DEMOQA_3").WebEdit("WebEdit").Set current_address
Browser("DEMOQA").Page("DEMOQA_3").WebEdit("Current Address").Set permanant_address
Browser("DEMOQA").Page("DEMOQA_3").WebButton("Submit").Click

End Function	
