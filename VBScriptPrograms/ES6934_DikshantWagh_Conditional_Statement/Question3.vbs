dim a
a=inputbox("Enter An Alphabet")
if len(a)=1 then  
  select case a
    case "A","E","I","O","U","a","e","i","o","u"
        msgbox "It's a VOWEL!!"
    case else
  	msgbox "It's a CONSONENT!!"
  end select
else
   msgbox "Enter Single Character"
end if 