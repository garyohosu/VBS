'http://www.atmarkit.co.jp/ait/articles/0407/08/news101_4.html 

Set objShell = WScript.CreateObject("WScript.Shell")
Set objExec = objShell.Exec("notepad.exe test.txt")
 ' �N����҂�
Do Until objShell.AppActivate(objExec.ProcessID)
   WScript.Sleep 1000
 Loop
 objShell.SendKeys "^f"      ' Ctrl+F�Ō����_�C�A���O�̕\��
 ' �_�C�A���O�̕\����҂�
 Do Until objShell.AppActivate("����")
   WScript.Sleep 100
 Loop
 objShell.SendKeys "WSH~"    ' WSH�Ɠ��͂���Enter�L�[������
 objShell.SendKeys "{ESC}"   ' �_�C�A���O�����
 WScript.Sleep 100
 objShell.AppActivate(objExec.ProcessID)
 objShell.SendKeys "{HOME}"  ' �s���ɖ߂�
 objShell.SendKeys "+{END}"  ' Shift+End �ōs���܂őI��
 objShell.SendKeys "^c"      ' Ctrl+C�ŃR�s�[
 objShell.SendKeys "%{F4}"   ' Alt+F4�ŏI��