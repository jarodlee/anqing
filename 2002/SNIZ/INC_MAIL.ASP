<%
'########## Snitz Forums 2000 Version 3.1 SR4 ####################
'#                                                               #
'#  汉化修改: 资源搜罗站                                         #
'#  电子邮件: cgier@21cn.com                                     #
'#  主页地址: http://www.sdsea.com                               #
'#            http://www.99ss.net                                #
'#            http://www.cdown.net                               #
'#	     http://www.wzdown.com                               #
'#	     http://www.13888.net                                #
'#  论坛地址:http://ubb.yesky.net                                #
'#  最后修改日期: 2001/03/12    中文版本：Version 3.1 SR4        #
'#################################################################
'# 原始来源                                                      #
'# Snitz Forums 2000 Version 3.1 SR4                             #
'# Copyright 2000 http://forum.snitz.com - All Rights Reserved   #
'#################################################################
'#【版权声明】                                                   #
'#                                                               #
'# 本软体为共享软体(shareware)提供个人网站免费使用，请勿非法修改,#
'# 转载，散播，或用于其他图利行为，并请勿删除版权声明。          #
'# 如果您的网站正式起用了这个脚本，请您通知我们，以便我们能够知晓#
'# 如果可能，请在您的网站做上我们的链接，希望能给予合作。谢谢！  #
'#################################################################
'# 请您尊重我们的劳动和版权，不要删除以上的版权声明部分，谢谢合作#
'# 如有任何问题请到我们的论坛告诉我们                            #
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
