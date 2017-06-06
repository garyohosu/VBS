strUser = "garyo"
strPass = "JCA03201"
 
Set Basp21 = CreateObject( "Basp21" )
 
strServer = "garyo.sakura.ne.jp"
nPort = 587
 
SMTPServer = strServer & ":" & nPort
MailTo = "あなたです <hantani@miyoshi.elec.co.jp>"
' CC
MailTo = MailTo & vbTab & "cc" & vbTab & "アカウント1 <xxxxxx@xxxxxx.jp>"
' BCC
MailTo = MailTo & vbTab & "bcc" & vbTab & "YAHOOアカウント <xxxxxxxx@yahoo.co.jp>"
 
MailFrom = "私です <garyo@garyo.sakura.ne.jp>"
MailFrom = MailFrom & vbTab & strUser & ":" & strPass
 
'strSubject = Wscript.Arguments(0)
'strTextbody = Replace( Wscript.Arguments(1), "\n", vbCrLf )
 
strSubject = "タイトル"
strTextbody = "本文"

 
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
	Wscript.Echo "メール送信が終了しました。"
end if
