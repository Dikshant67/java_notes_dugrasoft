'8.Use TimeSerial function returns the time for a specific hour, minute, and second.
'Using timeserial function we can create date using TimeSerial(hour,minute,second) and do arithmetic 
'operations inside the paranthesis
Option Explicit
dim date1,date2
date1=Timeserial(23,3,3)
date2=Timeserial(23-12,09,3-2)
msgbox date1
msgbox date2
