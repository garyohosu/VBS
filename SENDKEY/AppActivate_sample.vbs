'http://www.atmarkit.co.jp/ait/articles/0407/08/news101_4.html 

Set objShell = WScript.CreateObject("WScript.Shell")
Set objExec = objShell.Exec("notepad.exe test.txt")
 ' 起動を待つ
Do Until objShell.AppActivate(objExec.ProcessID)
   WScript.Sleep 1000
 Loop
 objShell.SendKeys "^f"      ' Ctrl+Fで検索ダイアログの表示
 ' ダイアログの表示を待つ
 Do Until objShell.AppActivate("検索")
   WScript.Sleep 100
 Loop
 objShell.SendKeys "WSH~"    ' WSHと入力してEnterキーを押す
 objShell.SendKeys "{ESC}"   ' ダイアログを閉じる
 WScript.Sleep 100
 objShell.AppActivate(objExec.ProcessID)
 objShell.SendKeys "{HOME}"  ' 行頭に戻る
 objShell.SendKeys "+{END}"  ' Shift+End で行末まで選択
 objShell.SendKeys "^c"      ' Ctrl+Cでコピー
 objShell.SendKeys "%{F4}"   ' Alt+F4で終了