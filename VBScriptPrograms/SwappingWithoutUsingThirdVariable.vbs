Dim first
Dim second
Dim third

msgbox "Enter two Numbers"
first= inputbox("Enter value for num1")
second= inputbox("Enter value for num2")
first=CInt(first)+CInt(second)
second=first-CInt(second)
first=first-second
msgbox "After Swapping  first "&first&" second "&second
msgbox first
msgbox second