option Explicit
on error resume Next
Dim stringval

stringval=Array("a","b")

msgbox stringval(2)

msgbox "Error Number = "&Err.Number
msgbox "Error Description = "&Err.Description	
Err.clear

msgbox "Error Number = "&Err.Number
msgbox "Error Description = "&Err.Description	
on error goto 0