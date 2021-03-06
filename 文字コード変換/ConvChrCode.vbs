'http://blog.goo.ne.jp/fukuriko/e/a40eb03d587d572f1389b24d17c60526

'【ＶＢＳｃｒｉｐｔによる文字コード変換】
'【使用例】
'  ' 変数str は Shift_JIS に変換された "あういうえお" が代入される
'  str = Encode("あいうえお","Shift_JIS")

Function Encode(strUni,cset) 
    Set objStm = CreateObject("ADODB.Stream")
    objStm.Open
    objStm.Type = 2
    objStm.Charset = cset
    objStm.WriteText strUni 
    objStm.Position = 0
    objStm.Type = 1
    Select Case UCase(cset)
      Case "UNICODE", "UTF-16"
        objStm.Position = 2
      Case "UTF-8"
        objStm.Position = 3
    End Select
    Encode = objStm.Read()


    objStm.Close
    Set objStm = Nothing
End Function

