'http://www.atmarkit.co.jp/ait/articles/0407/08/news101_4.html


Set objShell = WScript.CreateObject("WScript.Shell")
objShell.Run "notepad.exe"
WScript.Sleep 1000
objShell.SendKeys "hello"
