Option Explicit
Dim args, arg

Set args = WScript.Arguments
For Each arg In args
    Call MsgBox(arg,,"引数に入力した文字列")
Next
