'http://blog.goo.ne.jp/fukuriko/e/a40eb03d587d572f1389b24d17c60526

'�y�u�a�r�����������ɂ�镶���R�[�h�ϊ��z
'�y�g�p��z
'  ' �ϐ�str �� Shift_JIS �ɕϊ����ꂽ "������������" ����������
'  str = Encode("����������","Shift_JIS")

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

