'12.Give an example of FormatDateTime Function.
'formatdatetime function returns date in specific format
option Explicit
dim date1
date1="01-02-2022 22:16"
'd=CDate("2019-05-31 13:45")
msgbox FormatDateTime(date1)
msgbox FormatDateTime(date1,1)
msgbox FormatDateTime(date1,2)
msgbox FormatDateTime(date1,3)