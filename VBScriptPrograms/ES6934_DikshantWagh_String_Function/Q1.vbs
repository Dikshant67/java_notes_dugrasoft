dim a
str="ExpleoIndia"
strlen=len(str)
dim count

for i=1 to strlen step 1
  a=mid(str,i,1)
  select case a
    case "A","E","I","O","U","a","e","i","o","u"
       count=count+1
  end select
Next
msgbox count