dim str,str1
str="Expleo"

for i=len(str) to 1 step -1
 str1=str1+mid(str,i,1)
Next

msgbox "Before Reversing "&str
msgbox "After Reversing "&str1
