dim str,str1

str="ExpleoIndiaaaa"
for i=1 to len(str) step 1
	if instr(1,str1,mid(str,i,1))=0 then 
		str1=str1+mid(str,i,1)+","
	end if
Next
msgbox "Characters that Constitute "&str&" are "&str1    