'http://homepage2.nifty.com/nonnon/Chinamini/20100001/20100305.html
'��VBS��RS232C�̃T���v���\�[�X
' COM�|�[�g�ݒ�
Set MSComm1 = CreateObject("MSCommLib.MSComm")
MSComm1.CommPort = 20
MSComm1.Settings = "9600,N,8,1"

' COM�|�[�g�I�[�v��
MSComm1.PortOpen = True

' �f�[�^���M
MSComm1.Output = "12345����������"

' ACK�܂���NAK����M����܂ł܂�
Do
    WScript.Sleep(1000)
    RcvDat = RcvDat & MSComm1.Input
Loop While InStr(RcvDat, Chr(6)) <= 0 And _
           InStr(RcvDat, Chr(21)) <= 0

' ��M�f�[�^�\��
MsgBox RcvDat

' COM�|�[�g�N���[�Y
MSComm1.PortOpen = False