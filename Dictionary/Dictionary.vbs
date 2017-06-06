
Option Explicit
Dim objDictionary
'連想配列の作成
Set objDictionary = WScript.CreateObject("Scripting.Dictionary")
'キーと値の追加
objDictionary.Add "kona","Konata Izumi"
objDictionary.Add "kaga","Kagami Hiiragi"
objDictionary.Add "tuka","Tsukasa Hiiragi"
objDictionary.Add "yuki","Miyuki Takara"

'値の参照
MsgBox "キー「kaga」の値は、「" & objDictionary("kaga") & "」です。"

'キーの存在確認
If objDictionary.Exists("yuki") Then
	MsgBox "キー「yuki」は、存在します。"
End If

'要素の数を表示
MsgBox "Dictinaryに含まれる要素数：" & objDictionary.Count

Dim strKey, strMessage
'辞書の列挙
For Each strKey In objDictionary.Keys
        strMessage = strMessage & objDictionary(strKey) & vbCrLf
Next
MsgBox strMessage
objDictionary.RemoveAll() '辞書の全要素削除
objDictionary = Nothing
