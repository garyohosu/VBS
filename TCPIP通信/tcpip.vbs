' ----------------------------------------------------------- 
' TCP/IP�ŒʐM����T���v��(VBS/VBA)
' �ʐM�I�u�W�F�N�g����
' [regsvr32.exe NONCOMSCK.OCX]���K�v
Set Winsock1 = CreateObject("NonComSck.Winsock")

' TCP/IP�ڑ�
Winsock1.Connect "127.0.0.1", 8888

' TCP/IP�ڑ��҂�
Do While Winsock1.State = 6
    WScript.Sleep(500)
Loop

' �f�[�^���M(�������Byte�z��ɕϊ����đ��M)
Winsock1.SendData _
    Winsock1.StrToByteArray(Chr(2) & "1234567890" & Chr(3))

' �f�[�^��M
Winsock1.Start_EventForScript()
Do
    WScript.Sleep(500)
    Evt = Winsock1.GetEventParameters()
    If Ubound(Evt) >= 0 Then
        ' Evt(0) : �C�x���g��
        If Evt(0) = "DataArrival" Then
            ' Evt(9) : ��M�f�[�^��Byte�z��
            ' Byte�z��𕶎���ɕϊ�
            MsgBox Winsock1.ByteArrayToStr(Evt(9))
            Exit Do
        End If
    End If
Loop
Winsock1.End_EventForScript()

' TCP/IP�ؒf
Winsock1.Close2

MsgBox "�I��"
' ----------------------------------------------------------- 
