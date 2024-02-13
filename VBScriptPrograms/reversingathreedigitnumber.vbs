Dim first
Dim temp
dim temp2
'msgbox "Enter a Number"
'first= inputbox("Enter a three digit number")
first = 999
first=Cint(first)
'second= inputbox("Enter value for num2")
temp=first MOD 10
Msgbox temp
temp2=cint(cint(first-temp)/10) MOD 10
msgbox temp2
temp=temp*100

temp=temp+cint(temp2)*10
first=int(first/100)
msgbox(first)
temp=temp+first
temp=Cint(temp)
'msgbox "After Swapping  first "&first&" second "&second
msgbox temp
'msgbox second