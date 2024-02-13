'7.Give an example of DatePart
option Explicit
dim date1
dim yearofdate
dim weekofdate 
date1="01-02-2020"
yearofdate=Datepart("yyyy",date1)
weekofdate=Datepart("ww",date1)

msgbox yearofdate
msgbox weekofdate