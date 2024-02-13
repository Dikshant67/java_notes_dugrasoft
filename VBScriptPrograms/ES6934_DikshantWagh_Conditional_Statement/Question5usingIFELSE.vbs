dim unit
dim cost
dim temp
dim temp2
dim temp3
'msgbox "Enter Number of Units"
unit=int(inputbox("Enter Number of Units"))
if unit < 51 THEN
   cost=unit*0.5
   cost=cost*1.2
   msgbox "Cost is "&cost
elseif unit>50 AND unit<151 then
   temp=unit-50
   cost=temp*0.75+50*0.5
   cost=cost*1.2
   msgbox "Cost is "&cost
elseif unit>150 AND unit<251 then
   temp=unit-50
   temp2=temp-100
   'temp3=temp2-100
   cost=100*0.75+50*0.5+temp2*1.2
   cost=cost*1.2
   msgbox "Cost is "&cost
elseif unit>250 then
   temp=unit-50
   temp2=temp-100
   temp3=temp2-100
   cost=100*0.75+50*0.5+100*1.2+temp3*1.5
   cost=cost*1.2
   msgbox "Cost is "&cost
end if