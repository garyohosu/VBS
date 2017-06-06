
' ArrayListの作成
Dim myArrayList
Set myArrayList = CreateObject("System.Collections.ArrayList")

' 要素を追加していく
myArrayList.add "Alice"
myArrayList.add "Bob"
myArrayList.add "Charlie"

' For Each文で要素を表示
Dim elem
For Each elem In myArrayList
   Wscript.echo elem
Next

' For文で要素を表示
Dim i
Dim iMax : iMax = myArrayList.Count - 1
For i=0 To iMax
   Wscript.echo myArrayList(i)
Next

' 新しい要素を挿入
myArrayList.Insert 1, "Dave"
myArrayList.Insert 2, "Ellen"

' 要素をソート
myArrayList.Sort

' 特定の要素を削除
myArrayList.Remove "Dave"

' 通常の配列に変換
Dim myArray : myArray = myArrayList.ToArray()

' 全要素を削除
myArrayList.Clear
