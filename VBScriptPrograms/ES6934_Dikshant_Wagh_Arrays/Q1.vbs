dim arr,temp
arr=Array(34,23,36,98,19)
for i=0 to ubound(arr) 
 msgbox arr(i)
next
msgbox "Before Sorting"
for i=ubound(arr) to 0 step -1
 for j=0 to i step 1
   if arr(i)>arr(j) then
       temp=arr(i)
	arr(i)=arr(j)
	arr(j)=temp
   end if
  next
next		
msgbox "After Sorting"
for i=0 to ubound(arr) 
 msgbox arr(i)
next