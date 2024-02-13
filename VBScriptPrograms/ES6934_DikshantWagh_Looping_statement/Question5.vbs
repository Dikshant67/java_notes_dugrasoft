dim x
x=Array(35,44,99,66,98,76)
dim temp,temp2
temp=x(0)
for i=0 to ubound(x)
  if(x(i)>temp) then
     temp=x(i)
  end if
next

for i=0 to ubound(x)
  if(x(i) < temp and x(i)>temp2 ) then
     temp2=x(i)
  end if
next
msgbox "Second Highest Number is "&temp2

