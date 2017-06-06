
dim dic

set dic = CreateObject("Scripting.dictionary")

dic.add "A",1
dic.add "B",2

msgbox(dic.exists("A"))
msgbox(dic.exists("C"))

dim item

for each item in dic.keys
	msgbox(item & " " & dic(item))
next

