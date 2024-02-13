dim str,str1,count,count1
str="Expleo India InfoSystem"
for i=len(str) to 1 step -1
   count=count+1
   if asc(mid(str,i,1))=32 then
	str1=str1+mid(str,i,count)
	count=0  
   end if 
next
str1=str1+" "+left(str,count)
msgbox str1	