'http://www.atmarkit.co.jp/ait/articles/0707/26/news128.html

'cscript.exe //NoLogo ファイル名で起動すること

'http://homepage2.nifty.com/nihon-nouen/programming-stdinout.htm
'StdIn	Read	指定された文字数を入力ストリームから読み込み、結果の文字列を返す
'       ReadAll	入力ストリーム全体を読み込み、結果の文字列を返す
'       ReadLine	行全体 (改行文字の直前まで) を入力ストリームから読み込み、結果の文字列を返す
'       Skip	入力ストリームの読み込み中に、指定された文字数をスキップする
'       SkipLine	入力ストリームの読み込み中に、次の 1 行をスキップする
'StdOut	Write	指定された文字列を出力ストリームに書き込む
'       WriteBlankLines	指定された数の改行文字を出力ストリームに書き込む
'       WriteLine	指定された文字列と改行文字を出力ストリームに書き込む


Option Explicit
Dim objStdIn, objStdOut
Set objStdIn  = WScript.StdIn  '標準入力ストリームを返す
Set objStdOut = WScript.StdOut '標準出力ストリームを返す

Dim strFromStdIn
'標準入力から文字列を1行読み込む
strFromStdIn = objStdIn.ReadLine()

'標準出力に文字列を1行書き込む
objStdOut.WriteLine strFromStdIn

objStdIn.Close()  '標準入力ストリームを閉じる
objStdOut.Close() '標準出力ストリームを閉じる

'オブジェクトの破棄
Set objStdIn  = Nothing
Set objStdOut = Nothing

