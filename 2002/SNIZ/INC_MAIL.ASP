<%
'########## Snitz Forums 2000 Version 3.1 SR4 ####################
'#                                                               #
'#  �����޸�: ��Դ����վ                                         #
'#  �����ʼ�: cgier@21cn.com                                     #
'#  ��ҳ��ַ: http://www.sdsea.com                               #
'#            http://www.99ss.net                                #
'#            http://www.cdown.net                               #
'#	     http://www.wzdown.com                               #
'#	     http://www.13888.net                                #
'#  ��̳��ַ:http://ubb.yesky.net                                #
'#  ����޸�����: 2001/03/12    ���İ汾��Version 3.1 SR4        #
'#################################################################
'# ԭʼ��Դ                                                      #
'# Snitz Forums 2000 Version 3.1 SR4                             #
'# Copyright 2000 http://forum.snitz.com - All Rights Reserved   #
'#################################################################
'#����Ȩ������                                                   #
'#                                                               #
'# ������Ϊ��������(shareware)�ṩ������վ���ʹ�ã�����Ƿ��޸�,#
'# ת�أ�ɢ��������������ͼ����Ϊ��������ɾ����Ȩ������          #
'# ���������վ��ʽ����������ű�������֪ͨ���ǣ��Ա������ܹ�֪��#
'# ������ܣ�����������վ�������ǵ����ӣ�ϣ���ܸ��������лл��  #
'#################################################################
'# �����������ǵ��Ͷ��Ͱ�Ȩ����Ҫɾ�����ϵİ�Ȩ�������֣�лл����#
'# �����κ������뵽���ǵ���̳��������                            #
'#################################################################

select case lcase(strMailMode) 
	case "aspmail"
		Set objNewMail = Server.CreateObject("SMTPsvg.Mailer")
		objNewMail.FromName = strFromName
		objNewMail.FromAddress = strSender
		'objNewMail.AddReplyTo = strSender
		objNewMail.RemoteHost = strMailServer
		objNewMail.AddRecipient strRecipientsName, strRecipients
		objNewMail.Subject = strSubject
		objNewMail.BodyText = strMessage
		on error resume next '## Ignore Errors
		SendOk = objNewMail.SendMail
		If not(SendOk) <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & objNewMail.Response & "</li>"
		End if
	case "aspemail"
		Set objNewMail = Server.CreateObject("Persits.MailSender")
		objNewMail.FromName = strFromName
		objNewMail.From = strSender
		objNewMail.AddReplyTo strSender
		objNewMail.Host = strMailServer
		objNewMail.AddAddress strRecipients, strRecipientsName
		objNewMail.Subject = strSubject
		objNewMail.Body = strMessage
		on error resume next '## Ignore Errors
		objNewMail.Send
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if
	case "aspqmail"
		Set objNewMail = Server.CreateObject("SMTPsvg.Mailer")
		objNewMail.QMessage = 1
		objNewMail.FromName = strFromName
		objNewMail.FromAddress = strSender
		objNewMail.RemoteHost = strMailServer
		objNewMail.AddRecipient strRecipientsName, strRecipients
		objNewMail.Subject = strSubject
		objNewMail.BodyText = strMessage
		on error resume next '## Ignore Errors
		objNewMail.SendMail
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if
	case "cdonts"
		Set objNewMail = Server.CreateObject ("CDONTS.NewMail")
		objNewMail.BodyFormat = 1
		objNewMail.MailFormat = 0
		on error resume next '## Ignore Errors
		objNewMail.Send strSender, strRecipients, strSubject, strMessage
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if
		on error resume next '## Ignore Errors
	case "chilicdonts"
		Set objNewMail = Server.CreateObject ("CDONTS.NewMail")
		on error resume next '## Ignore Errors
		objNewMail.Send strSender, strRecipients, strSubject, strMessage
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if
		on error resume next '## Ignore Errors
	case "dkqmail"
		Set objNewMail = Server.CreateObject("dkQmail.Qmail")
		objNewMail.FromEmail = strSender
		objNewMail.ToEmail = strRecipients
		objNewMail.Subject = strSubject
		objNewMail.Body = strMessage
		objNewMail.CC = ""
		objNewMail.MessageType = "TEXT"
		on error resume next '## Ignore Errors
		objNewMail.SendMail()
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if
	case "geocel"
		set objNewMail = Server.CreateObject("Geocel.Mailer")
		objNewMail.AddServer strMailServer, 25
		objNewMail.AddRecipient strRecipients, strRecipientsName
		objNewMail.FromName = strFromName
		objNewMail.FromAddress = strFrom
		objNewMail.Subject = strSubject
		objNewMail.Body = strMessage
		on error resume next '##  Ignore Errors
		objNewMail.Send()
		if Err <> 0 then 
			Response.Write "Your request was not sent due to the following error: " & Err.Description 
		else
			Response.Write "Your mail has been sent..."
		end if
	case "iismail"
		Set objNewMail = Server.CreateObject("iismail.iismail.1")
		MailServer = strMailServer
		objNewMail.Server = strMailServer
		objNewMail.addRecipient(strRecipients)
		objNewMail.From = strSender
		objNewMail.Subject = strSubject
		objNewMail.body = strMessage
		on error resume next '## Ignore Errors
		objNewMail.Send
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if
	case "jmail"
		Set objNewMail = Server.CreateObject("Jmail.smtpmail")
		objNewMail.ServerAddress = strMailServer
		objNewMail.AddRecipient strRecipients
		objNewMail.Sender = strSender
		objNewMail.Subject = strSubject
		objNewMail.body = strMessage
		objNewMail.priority = 3
		on error resume next '## Ignore Errors
		objNewMail.execute
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if
	case "smtp"
'# 		Set objNewMail = Server.CreateObject("SmtpMail.SmtpMail.1")
'# 		objNewMail.MailServer = strMailServer
'# 		objNewMail.Recipients = strRecipients
'# 		objNewMail.Sender = strSender
'# 		objNewMail.Subject = strSubject
'# 		objNewMail.Message = strMessage
'# 		on error resume next '## Ignore Errors
'# 		objNewMail.SendMail2
		If Err <> 0 Then 
			Err_Msg = Err_Msg & "<li>Your request was not sent due to the following error: " & Err.Description & "</li>"
		End if
end select

Set objNewMail = Nothing
%>