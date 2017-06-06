'正規表現を使ってパターンに一致するか調べる

'Newを使ってRegExpオブジェクトを取得する。パターンはRegExpオブジェクトのPatternプロパティに代入する。

Dim re
Set re = New RegExp
re.Pattern = "^A.*$"
If re.Test("ABC") Then
    WScript.Echo "matched"
End If
'大文字小文字を無視する場合はIgnoreCaseにTrueを代入する。

re.IgnoreCase = True
'最初の一致のみではなく、文字列全体を検索対象にする場合にはGlobalにTrueを代入する。

re.Global = True
'正規表現を使って一致する文字列を取得する

'RegExpオブジェクトのExecuteを使う。

Dim re, matches
Set re = New RegExp
re.Pattern = "^A.*$"
Set matches = re.Execute("ABC")
If matches.Count > 0 Then
    WScript.Echo matches(0)
End If
'出力結果は「ABC」となる。

'正規表現を使ってキャプチャした文字列を取得する

'Itemプロパティを利用してSubMatchesを取得する。

Dim re, matches
Set re = New RegExp
re.Pattern = "^([A-F]+)_([A-F]+).+"
Set matches = re.Execute("ABC_DEF_GHI")
If matches.Count > 0 Then
    WScript.Echo matches.Item(0).SubMatches.Item(0)
    WScript.Echo matches.Item(0).SubMatches.Item(1)
End If
'出力結果は「ABC」「DEF」となる。

'正規表現を使って文字列を置換する

'RegExpオブジェクトのReplaceを使う。

Dim re, matches
Set re = New RegExp
re.Pattern = "^ABC_"
WScript.Echo re.Replace("ABC_DEF_GHI", "XXX_")
'出力結果は「XXX_DEF_GHI」となる。

'正規表現を使って文字列を置換する（後方参照）

'RegExpオブジェクトのReplaceを使って置換文字列中に$数字の形式で参照する。

Dim re, matches
Set re = New RegExp
re.Pattern = "^([A-Z]{2})(.*)"
WScript.Echo re.Replace("ABCDE", "XX$2")
'出力結果は「XXCDE」となる。

