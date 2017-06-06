'5. Dictionary オブジェクト
'VBScript にはハッシュ変数はありませんが、Scripting.Dictionary オブジェクトを用いることで代替えが可能です。
'Dictionary オブジェクトの宣言と代入

Dim hash

Set hash = CreateObject("Scripting.Dictionary")
hash.Add "a", 1
hash.Add "b", 2

'Dictionary の要素の参照と代入
' 要素の参照
WScript.Echo hash("a") ' 1
WScript.Echo hash("b") ' 2

' 要素の代入
hash("a") = 5
hash("b") = 7

'Dictionary のプロパティとメソッド
' キーの取得
keys = hash.Keys

' 値の取得
values = hash.Items

' キーの存在確認
hash.Exists("a")

' キーの削除
hash.Remove "a"
