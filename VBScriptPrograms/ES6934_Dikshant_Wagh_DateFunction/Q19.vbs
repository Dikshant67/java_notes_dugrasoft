'19.How can we display the week of the year and quarter of the year.
Option Explicit
dim quarter_name
dim week_name
week_name=Datepart("ww","06-7-2000")
quarter_name=Datepart("q","06-7-2000")
msgbox "06-7-2000 Week "&week_name 
msgbox "06-7-2000 Quarter "&quarter_name