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
<table border="0" width="100%">
  <tr>
    <td width="33%" align="left" nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    <img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="default.asp">返回论坛首页</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="admin_home.asp">论坛管理中心</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" height=15 width=15 border="0">&nbsp;邮件发送设定<br>
    </font></td>
  </tr>
</table>
<%
	if Request.Form("Method_Type") = "Write_Configuration" then 
		Err_Msg = ""
		if Request.Form("strMailServer") = "" and Request.Form("strEmail") = "1" then 
			Err_Msg = Err_Msg & "<li>你必须输入邮件主机位址！</li>"
		end if
		if ((lcase(left(Request.Form("strMailServer"), 7)) = "http://") or (lcase(left(Request.Form("strMailServer"), 8)) = "https://") or Request.Form("strMailServer") = "") and Request.Form("strEmail") = "1" then
			Err_Msg = Err_Msg & "<li>不要在邮件主机位址前加上 <b>http://</b>, <b>https://</b> or <b>file://</b></li>"
		end if
		if Request.Form("strSender") = "" then 
			Err_Msg = Err_Msg & "<li>你必须输入管理员的电子邮件位址</li>"
		else
			if EmailField(Request.Form("strSender")) = 0 and Request.Form("strSender") <> "" then 
				Err_Msg = Err_Msg & "<li>你必须输入合法的管理员电子邮件地址</li>"
			end if
		end if
	
		if Err_Msg = "" then

			'## Forum_SQL
			strSql = "UPDATE " & strTablePrefix & "CONFIG "
			strSql = strSql & " SET C_STREMAIL                 = " & Request.Form("strEmail") & ""
			strSql = strSql & ",    C_STRMAILMODE              = '" & Request.Form("strMailMode") & "'"
			if Request.Form("strMailServer") <> "" then
				strSql = strSql & ",    C_STRMAILSERVER            = '" & Request.Form("strMailServer") & "'"
			end if
			if Request.Form("strSender") <> "" then
				strSql = strSql & ",    C_STRSENDER                = '" & Request.Form("strSender") & "'"
			end if
			strSql = strSql & ",    C_STRUNIQUEEMAIL           = " & Request.Form("strUniqueEmail") & ""
			strSql = strSql & ",    C_STRLOGONFORMAIL          = " & Request.Form("strLogonForMail") & ""
		
			my_Conn.Execute(strSql)
			Application(strCookieURL & "ConfigLoaded") = ""

%>
<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">邮件发送设定已经更新完毕！</font></p>
<meta http-equiv="Refresh" content="2; URL=admin_home.asp">

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">修改完成！</font></p>

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="admin_home.asp">回到管理中心</font></a></p>
<%		else %>
	<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">你输入的资料有问题或没有填写完整</font></p>

	<table align=center border=0>
	  <tr>
	    <td><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
		<ul>
		<% =Err_Msg %>
		</ul>
	    </font></td>
	  </tr>
	</table>

	<p align=center><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>><a href="JavaScript:history.go(-1)">请返回重新输入</a></font></p>
<%		end if %>
<%	else %>
<form action="admin_config_email.asp" method="post" id="Form1" name="Form1">
<input type="hidden" name="Method_Type" value="Write_Configuration">
<table border="0" cellspacing="1" cellpadding="4" align=center bgcolor="<% =strPopUpBorderColor %>">
  <tr valign="top">
    <td bgcolor="<% =strHeadCellColor %>" colspan="2"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>"><b>邮件发送设定</b></font></td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">邮件服务器模式：</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    开启：<input type="radio" class="radio" name="strEmail" value="1" <% if lcase(strEmail) <> "0" then Response.Write("checked") %>> 
    关闭：<input type="radio" class="radio" name="strEmail" value="0" <% if lcase(strEmail) = "0" then Response.Write("checked") %>>
    <select name="strMailMode">
      <option value="aspmail"<% if (lcase(strMailMode)="aspmail" or lcase(MailMode)="aspmail") then Response.Write(" selected") %>>ASPMail</option>
      <option value="aspemail"<% if (lcase(strMailMode)="aspemail" or lcase(MailMode)="aspemail") then Response.Write(" selected") %>>ASPEMail</option>
      <option value="aspqmail"<% if (lcase(strMailMode)="aspqmail" or lcase(MailMode)="aspqmail") then Response.Write(" selected") %>>ASPQMail</option>
      <option value="cdonts"<% if (lcase(strMailMode)="cdonts" or lcase(MailMode)="cdonts") then Response.Write(" selected") %>>CDONTS</option>
      <option value="chilicdonts"<% if (lcase(strMailMode)="chilicdonts" or lcase(MailMode)="chilicdonts") then Response.Write(" selected") %>>Chili!Mail</option>
      <option value="dkqmail"<% if (lcase(strMailMode)="dkqmail") then Response.Write(" selected") %>>dkQMail</option>
      <option value="geocel"<% if (lcase(strMailMode)="geocel" or lcase(MailMode)="geocel") then Response.Write(" selected") %>>GeoCel</option>
      <option value="iismail"<% if (lcase(strMailMode)="iismail" or lcase(MailMode)="iismail") then Response.Write(" selected") %>>IISMail</option>
      <option value="jmail"<% if (lcase(strMailMode)="jmail" or lcase(MailMode)="jmail") then Response.Write(" selected") %>>JMail</option>
      <option value="smtp"<% if (lcase(strMailMode)="smtp" or lcase(MailMode)="smtp") then Response.Write(" selected") %>>SMTP</option>
    </select>
    <a href="JavaScript:openWindow3('pop_config_help.asp#email')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </font></td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">邮件服务器地址：</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    <input type="text" name="strMailServer" size="25" value="<% =strMailServer %>">
    <a href="JavaScript:openWindow3('pop_config_help.asp#mailserver')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">论坛管理员邮件地址：</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    <input type="text" name="strSender" size="25" value="<% =strSender %>">
    <a href="JavaScript:openWindow3('pop_config_help.asp#sender')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">邮件地址不得重复注册：</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    开启：<input type="radio" class="radio" name="strUniqueEmail" value="1" <% if strUniqueEmail = "1" then Response.Write("checked") %>> 
    关闭：<input type="radio" class="radio" name="strUniqueEmail" value="0" <% if strUniqueEmail <> "1" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#UniqueEmail')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </font></td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">发送电子邮件：</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    开启：<input type="radio" class="radio" name="strLogonForMail" value="1" <% if strLogonForMail = "1" then Response.Write("checked") %>> 
    关闭：<input type="radio" class="radio" name="strLogonForMail" value="0" <% if strLogonForMail <> "1" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#LogonForMail')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </font></td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" colspan="2" align="center"><input type="submit" value="提交设置" id="submit1" name="submit1"> <input type="reset" value="重新设置" id="reset1" name="reset1"></td>
  </tr>
</table>
</form>
</font>
<%	end if %>
<!--#INCLUDE file="inc_footer.asp" -->
<% Else %>
<% Response.Redirect "admin_login.asp" %>
<% End IF %>
