Dim mar,eng,maths,hist,sci,sum,per
MsgBox "Enter Marks of the Five Subjects"
mar=int(inputbox("Enter Marks for Marathi"))
eng=int(inputbox("Enter Marks for English"))
maths=int(inputbox("Enter Marks for Mathematics"))
hist=int(inputbox("Enter Marks for History"))
sci=int(inputbox("Enter Marks for Science"))
sum=mar+eng+maths+hist+sci
per=sum/500
per=per*100
MsgBox "Percentage is "&per

if per >= 90 then
 msgbox "Grade A"
elseif per >=80 then
 msgbox "Grade B"
elseif per >= 70 then 
 msgbox "Grade C"
elseif per >= 60 then 
 msgbox "Grade D"
elseif per >= 40 then 
 msgbox "Grade E"
elseif per < 40 then 
 msgbox "Grade F"
end if
