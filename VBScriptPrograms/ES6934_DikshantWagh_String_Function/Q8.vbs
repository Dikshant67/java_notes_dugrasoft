dim str
str="ExpleoIndia"
str1=str+"&"
dim count,count1,i
i=1
count=0
while (mid(str1,i,1)<>"&")
  count=count+1
  i=i+1	
Wend	
msgbox "Length of the given String is "&count