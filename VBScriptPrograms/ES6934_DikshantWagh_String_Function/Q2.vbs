dim a
str="ExPlEoInDiA"
strlen=len(str)
dim ucount,lcount
dim lowercase,uppercase

for i=1 to strlen step 1
   a=mid(str,i,1)
   If asc(a)>=65 AND asc(a)<=90 then
       ucount=ucount+1
       lowercase=lowercase+a+","
   elseif asc(a)>97 AND asc(a)<123 then
	lcount=lcount+1	
        uppercase=uppercase+a+","
   else
	msgbox "Invalid Input"
   end If	
Next
msgbox uppercase&" Uppercase Letter Count "&ucount&vbnewline&lowercase&" Lowercase Letter Count "&lcount