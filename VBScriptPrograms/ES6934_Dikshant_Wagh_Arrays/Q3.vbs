option Explicit
dim arrFruits:arrFruits=Array("Apple","Grapes","Guava","Blueberries","Banana")
dim fruitNames
dim fruit

fruitNames=Filter(arrFruits,"B")

for Each fruit in fruitNames
    msgbox fruit
 Next

