strUser = "garyo"
strPass = "JCA03201"
 
Set Basp21 = CreateObject( "Basp21" )
 
strServer = "garyo.sakura.ne.jp"
nPort = 587
 
SMTPServer = strServer & ":" & nPort
MailTo = "���Ȃ��ł� <hantani@miyoshi.elec.co.jp>"
' CC
MailTo = MailTo & vbTab & "cc" & vbTab & "�A�J�E���g1 <xxxxxx@xxxxxx.jp>"
' BCC
MailTo = MailTo & vbTab & "bcc" & vbTab & "YAHOO�A�J�E���g <xxxxxxxx@yahoo.co.jp>"
 
MailFrom = "���ł� <garyo@garyo.sakura.ne.jp>"
MailFrom = MailFrom & vbTab & strUser & ":" & strPass
 
'strSubject = Wscript.Arguments(0)
'strTextbody = Replace( Wscript.Arguments(1), "\n", vbCrLf )
 
strSubject = "�^�C�g��"
strTextbody = "�{��"

 
ErrMessage = Basp21.SendMail( _
SMTPServer, _
MailTo, _
MailFrom, _
strSubject, _
strTextbody, _
"" _
)
if ErrMessage <> "" then
	Wscript.Echo ErrMessage
else
	Wscript.Echo "���[�����M���I�����܂����B"
end if
