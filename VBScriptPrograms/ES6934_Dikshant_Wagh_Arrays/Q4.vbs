Option Explicit
'Given String
dim str
dim fruitnames,filteredfruitname
dim fruit
str="Apple~Grapes~Banana~Guava~Blueberries"
fruitnames=Split(str,"~")
filteredfruitname=Filter(fruitnames,"G")
for each fruit in filteredfruitname
 msgbox fruit
 Next