'17.Difference between weekday and weekdayName?
'weekday returns the weekday number in integer format 
'weekdayname returns the weekday name in string format
Option Explicit
dim weekday1
dim weekdayname1
weekday1=weekday("06-7-2000")
weekdayname1=weekdayname(weekday("06-7-2000"))
msgbox weekday1&" "&weekdayname1