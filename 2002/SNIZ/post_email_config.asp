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
%>
<!--#INCLUDE FILE="config.asp" -->
<% If Session(strCookieURL & "Approval") = "15916941253" Then %>
<!--#INCLUDE file="inc_functions.asp" -->
<!--#INCLUDE file="inc_top.asp" -->
<%
	strEmailMessage = ""
	str_m_type = "newreply"
	strEmailSubject = request.form("tsubject")
	strEmailMessage = request.form("tmessage")
	if strEmailMessage <> "" then
		strsql = "Update " & strTablePrefix & "email_config Set "
	    strsql = strsql & "m_subject='" & strEmailSubject & "', "
	    strsql = strsql & "m_message='" & strEmailMessage & "' "
	    strsql = strsql & "Where m_type='" & str_m_type & "'"
	    my_conn.execute(strsql)
	end if
	strEmailMessage = ""
	str_m_type = "replynot"
	strEmailSubject = request.form("rsubject")
	strEmailMessage = request.form("rmessage")
	if strEmailMessage <> "" then
		strsql = "Update " & strTablePrefix & "email_config Set "
	    strsql = strsql & "m_subject='" & strEmailSubject & "', "
	    strsql = strsql & "m_message='" & strEmailMessage & "' "
	    strsql = strsql & "Where m_type='" & str_m_type & "'"
	    my_conn.execute(strsql)
	end if

    
    response.write("<p><center>邮件通知信息已经修改完毕, <a href=""admin_email_messages.asp"">点击这里返回</a></center><p>")


%>
<!--#INCLUDE file="inc_footer.asp" -->
<% Else %>
<% Response.Redirect "admin_login.asp" %>
<% End IF %>

