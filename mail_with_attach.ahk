;#NoTrayIcon
#Persistent
#SingleInstance force




!m::
        try
		{
		;Mail(from,to,subject,content,attach*)
		attach := ["E:\OOP\mail_ahk\1.txt","E:\OOP\mail_ahk\_tmp_old","E:\OOP\mail_ahk\1.rar"]  ;add attach
		Mail("wy<from_mail@163.com>","to_mail@163.com",A_MM "-" A_DD "  " A_Hour ":" A_Min,"12345",attach*)
		MsgBox,,,you have a mail!,6
	    }
		catch
        {
         MsgBox, 16,, There was a problem while mail to 163.com!,5
		}
return



Mail(from,to,subject,content,attach*){
cdoSendUsingPort:=2	
cdoBasic:=1 
NameSpace := "http://schemas.microsoft.com/cdo/configuration/"
Email := ComObjCreate("CDO.Message")
Email.From := from
Email.To := to
Email.Subject := subject
Email.Htmlbody := content
;~ Email.Textbody := content
for k,v in attach
{
	IfExist, % v
	{
		Email.AddAttachment(v)
	}
}
Email.Configuration.Fields.Item(NameSpace "sendusing") :=2
Email.Configuration.Fields.Item(NameSpace "smtpserver") := "smtp.163.com" ;SMTP∑˛ŒÒ∆˜µÿ÷∑
Email.Configuration.Fields.Item(NameSpace "smtpserverport") := 465
Email.Configuration.Fields.Item(NameSpace "connectiontimeout)") := 10
Email.Configuration.Fields.Item(NameSpace "smtpauthenticate") :=1
Email.Configuration.Fields.Item(NameSpace "smtpusessl") :=true
Email.Configuration.Fields.Item(NameSpace "sendusername") := "from_mail@163.com" ;” œ‰’À∫≈
Email.Configuration.Fields.Item(NameSpace "sendpassword") := "passwd" ;” œ‰ ⁄»®¬Î
Email.Configuration.Fields.update
Email.Send
}


GuiClose:
ExitApp