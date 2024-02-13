option explicit

dim sizeofFirstArray,sizeofSecondArray,sizeofthirdArray
dim firstArray,secondArray,thirdArray



MsgBox "Enter Sizes and Elements of Arrays"

sizeofFirstArray=int(inputbox("Enter Size of First Array"))

redim firstArray(sizeofFirstArray-1)
msgbox "Enter Elements of First Array"
for i=0 to ubound(firstArray)
   firstArray(i)=inputbox("Enter "&i+1&" Element of First Array")
next

sizeofSecondArray=int(inputbox("Enter Size of Second Array"))

redim secondArray(sizeofSecondArray-1)
msgbox "Enter Elements of Second Array"
for i=0 to ubound(secondArray) 
   secondArray(i)=inputbox("Enter "&i+1&" Element of Second Array")
next

dim i,j,k


sizeofthirdArray=ubound(firstArray)+ubound(secondArray)+1
redim thirdArray(sizeofthirdArray)
for i=0 to ubound(firstArray) 
   thirdArray(i)=firstArray(i)
next

for j=0 to ubound(secondArray)
    thirdArray(i)=secondArray(j)
	i=i+1
next
msgbox "Third Concated Array Elements Loop Starts"
for k=0 to ubound(thirdArray)
  MsgBox thirdArray(k)
next  
msgbox "Third Concated Array Elements Loop Ends"
  

