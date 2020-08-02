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
<!--#INCLUDE FILE="inc_functions.asp" --> 
<!--#INCLUDE FILE="inc_top_short.asp" -->
<%
if Request.QueryString("mode") = "DoIt" then
	Err_Msg = ""
	if (Request.Form("Name") = "") then 
		Err_Msg = Err_Msg & "<li>你必须输入你的用户名</li>"
	end if
	if (Request.Form("Email") = "") then 
		Err_Msg = Err_Msg & "<li>你必须输入你的电子邮件地址！</li>"
	end if
	if (EmailField(Request.Form("Email")) = 0) then 
		Err_Msg = Err_Msg & "<li>必须填写正确的电子邮件地址</li>"
	end if
	'##  Emails Password to Requestor if Authenticated Properly.  
	'##  This needs to be Edited to use your own email component
	'##  if you don't have one, try the w3Jmail component from www.dimac.net it's free!
	if lcase(strEmail) = "1" then
		if (Err_Msg = "") then
			'## Forum_SQL
			strSql = "SELECT M_NAME, M_PASSWORD, M_EMAIL "
			strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
			strSql = strSql & " WHERE M_NAME = '" & Request.Form("Name") & "' AND "
			strSql = strSql & "       M_EMAIL = '" & Request.Form("email") & "'"

			set rs = my_Conn.Execute (strSql)

			strRecipientsName = Request.Form("Name")
			strRecipients = request.form("email")
			strSubject = strForumTitle & " " & "密码寻回"
			strMessage = "由于你在 " & strForumTitle & " 的注册密码丢失，现在我们通过email重新发送密码给你，请妥善保管好..我们的网站地址： " & strForumURL & vbCrLf & vbCrLf

			if rs.EOF or rs.BOF then
				strMessage = strMessage & "Sorry. The details you supplied were wrong or unknown." & vbCrLf & vbCrLf
				strMessage = strMessage & "Please register again at " & strForumURL & "/register.asp?mode=Register" & vbCrLf
			Else
				strMessage = strMessage & " 你的密码是: " & rs("M_PASSWORD") & vbCrLf
			end if
			
			strMessage = strMessage & "欢迎光临 " & strForumTitle
%>
<!--#INCLUDE FILE="inc_mail.asp" -->
<%
			rs.close
			set rs = nothing
'			Response.Write vbCrLf & "<br>发件人: " & strSender & vbCrLf & "<br>Recipients: " & strRecipients & vbCrLf & "<br>主题: " & strSubject & vbCrLf & "<br>Message: " & strMessage & vbCrLf & "<br>"
			Response.Write "<p><font face=" & strDefaultFontFace & " size=" & strHeaderFontSize & ">密码已经发送回你的信箱！</font></p>" & vbCrLf
			Response.Write "<p><font face=" & strDefaultFontFace & " size=" & strDefaultFontSize & ">请检查你的电子邮箱.</font></p>" & vbCrLf
		else
%>
	<p><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">这封电子邮件有问题，可能有输入错误或没有填写完整。</font></p>

	<table>
	  <tr>
	    <td>
		<ul>
		<% =Err_Msg %>
		</ul>
	    </td>
	  </tr>
	</table>

	<p><font size="<% =strDefaultFontSize %>"><a href="JavaScript:history.go(-1)">请返回重新输入</a></font></p>

	<p><font size="<% =strDefaultFontSize %>"><a href="JavaScript:onClick= window.close()">关闭窗口</A></font></p>
<%
			Response.End 
		end if
	end if
Else 
%>
<p><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">密码寻回</font></p>

<form action="pop_pword.asp?mode=DoIt" method="post">
<table border="0" width="75%" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="<% =strPopUpBorderColor %>">
    <table border="0" width="100%" cellspacing="1" cellpadding="1">
    <TBODY>
      <TR>
        <TD bgColor=<% =strPopUpTableColor %> align=right nowrap><b><FONT face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>>用户名称：</font></b></TD>
        <TD bgColor=<% =strPopUpTableColor %>><INPUT name=Name size=25></TD>
      <TR>
        <TD bgColor=<% =strPopUpTableColor %> align=right nowrap><b><FONT face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>>邮件地址:</font></b></TD>
        <TD bgColor=<% =strPopUpTableColor %>><INPUT name=email size=25 value=""></TD>
      </TR>
        <TD bgColor=<% =strPopUpTableColor %> align=middle colSpan=2><INPUT name=submit1 type=submit value=提交></TD>
      </TR>
    </TBODY>
    </TABLE>
    </td>
  </tr>
</table>
</form>
<% end if %>
<!--#INCLUDE FILE="inc_footer_short.asp" -->
