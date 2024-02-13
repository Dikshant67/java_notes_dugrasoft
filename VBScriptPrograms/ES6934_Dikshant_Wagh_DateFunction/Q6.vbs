'6.How can we calculate the Date Difference
option Explicit
dim fromdate 
dim todate 
fromdate = "01-01-2020 09:00:00"
todate = "029-08-2023 00:00:00"
msgbox DateDiff("yyyy",fromdate,todate)
msgbox DateDiff("m",fromdate,todate)
msgbox DateDiff("d",fromdate,todate)