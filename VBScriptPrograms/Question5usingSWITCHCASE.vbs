dim unit
dim cost
dim temp
dim temp2
dim temp3
dim var
'msgbox "Enter Number of Units"
unit=int(inputbox("Enter Number of Units"))
if unit < 51 THEN
  var=1
elseif unit>50 AND unit<151 then
  var=2
elseif unit>150 AND unit<251 then
  var=3
elseif unit>250 then
  var=5
end if

select case var
  case 1  
   cost=unit*0.5
   cost=cost*1.2
   msgbox "Cost is "&cost
  case 2
   temp=unit-50
   cost=temp*0.75+50*0.5
   cost=cost*1.2
   msgbox "Cost is "&cost
  case 3
   temp=unit-50
   temp2=temp-100
   temp3=temp2-100
   cost=100*0.75+50*0.5+100*1.2+temp3*1.5
   cost=cost*1.2
   msgbox "Cost is "&cost
  case 4
   temp=unit-50
   temp2=temp-100
   temp3=temp2-100
   cost=100*0.75+50*0.5+100*1.2+temp3*1.5
   cost=cost*1.2
   msgbox "Cost is "&cost
end select
