'18.How can we calculate the hour difference?
option Explicit
dim fromdate,todate
fromdate="08-09-2014 01:11:19"
todate= Now


msgbox DateDiff("h",fromdate,todate)