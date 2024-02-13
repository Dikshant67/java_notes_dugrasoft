dim str,count
str="if + if = 2 if"

for i=1 to len(str)-1 step 1
   if mid(str,i,1)="i" and mid(str,i+1,1)="f" then 
	count=count+1
   end if 	
Next

msgbox "Number of 'if' in the given statement "&count

