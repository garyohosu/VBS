Option Explicit
Dim args, arg

Set args = WScript.Arguments
For Each arg In args
    Call MsgBox(arg,,"�����ɓ��͂���������")
Next
