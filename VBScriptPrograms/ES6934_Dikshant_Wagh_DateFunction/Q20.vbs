'20. Display all Month and quarter of the year from the given Date.
option Explicit
dim date1
dim month1
dim quarter1
dim i,j
date1="06-07-2024"

month1=DatePart("m",date1)
quarter1=Datepart("q",date1)
for i=month1 to 12 step 1
   msgbox monthname(i)
 Next
for j=quarter1 to 4 step 1
  msgbox j   
next