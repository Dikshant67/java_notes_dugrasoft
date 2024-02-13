dim x,temp
x=int(inputbox("Enter a Number "))
temp=x
dim sum1,rem1
while temp>0
  rem1=temp MOD 10
  sum1=sum1+rem1
  temp=int(temp/10)
Wend
msgbox "Sum of the digits of Number "&x&" is "&sum1			