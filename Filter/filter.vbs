Dim aryIndex
Dim aryStrings
aryStrings = Array("山田", "佐藤", "中山", "山本", "鈴木")
aryIndex = Filter(aryStrings, "山", True, 1)     '"山田", "中山", "山本" の３要素がセットされます。
msgbox(Join(aryIndex, "/"))
aryIndex = Filter(aryStrings, "山", False, 1)    '"佐藤", "鈴木" の２要素がセットされます。
msgbox(Join(aryIndex , "/"))
