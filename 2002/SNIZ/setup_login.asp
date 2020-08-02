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

blnSetup = Request.Form("setup")
%>
<!--#INCLUDE FILE="config.asp" -->
<!--#INCLUDE FILE="inc_functions.asp" -->
<html>

<head>
<title>Forum-Setup Page</title>
<meta name="copyright" content="资源搜罗站(www.99ss.net)">
<Style><!--
a:link    {color:<% =strLinkColor %>;text-decoration:<% =strLinkTextDecoration %>}
a:visited {color:<% =strVisitedLinkColor %>;text-decoration:<% =strVisitedTextDecoration %>}
a:hover   {color:<% =strHoverFontColor %>;text-decoration:<% =strHoverTextDecoration %>}
--></style>
</head>

<body bgColor="white" text="midnightblue" link="darkblue" aLink=red vLink="red" onLoad="window.focus()">
<%
set my_Conn = Server.CreateObject("ADODB.Connection")
my_Conn.Open strConnString

Name = Request.Form("Name")
Password = Request.Form("Password")

RequestMethod = Request.ServerVariables("Request_method")

if RequestMethod = "POST" Then

	'## Forum_SQL
	strSql = "SELECT COUNT(*) AS ApprovalCode "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE M_NAME = '" & Name & "' AND "
	strSql = strSql & "       M_PASSWORD = '" & Password & "' AND "
	strSql = strSql & "       M_LEVEL = 3"
	
	Set dbRs = my_Conn.Execute(strSql)

	If dbRS.Fields("ApprovalCode") = "1"  and ChkQuoteOk(Name) and ChkQuoteOk(Password) Then 
%>
<p>&nbsp;</p>
<p align="center"><font face="宋体, Arial, Helvetica" size="4">登陆成功！</font></p>
<% Session(strCookieURL & "Approval") = "15916941253" %>
<p>&nbsp;</p>
<p align="center"><p align="center"><font face="宋体, Arial, Helvetica" size="2"><a href="setup.asp?RC=3" target="_top">Click here to Continue.</a></font></p>
<meta http-equiv="Refresh" content="0; URL=setup.asp?RC=3">
<% Response.End %>
<% else %>
<div align=center><center>
	<p><font face="宋体, Arial, Helvetica" size="4">There has been a problem !</font></p>
</center></div>

<form action="setup_login.asp" method="post" id=Form1 name=Form1>
<input type="hidden" name="setup" value="Y">
<table width="50%" height="50%" align="center" border="0" cellspacing="0" cellpadding="5">
		<tr>
			<td bgColor=navyblue align="center"><p align="center"><font face="宋体, Arial, Helvetica" size="2"><b>你无权进入</b></font></p></td>
		</tr>
		<tr>
			<td bgColor=navyblue align="left"><p><font face="宋体, Arial, Helvetica" size="2">I如果你认为这个讯息是错误的，请重新尝试一遍</font></p></td>
		</tr>
		<tr>
			<td>
			<table border="0" cellspacing="0" cellpadding="5" border="0" cellspacing="2" cellpadding="0" align="center">
			<tr>
			  <td align="center" colspan="2" bgColor=navyblue><b><font face="宋体, Arial, Helvetica" size="2">管理员登陆</font></b></td>
			</tr>
			<tr>
			  <td align="right" nowrap><b><font face="宋体, Arial, Helvetica" size="2">名称：</font></b></td>
			  <td><input type="text" name="Name"></td>
			</tr>
			<tr>
			  <td align="right" nowrap nowrap><b><b><font face="宋体, Arial, Helvetica" size="2">密码：</font></b></td>
			  <td><input type="Password" name="Password"></td>
			</tr>
			<tr>
			  <td colspan="2" align="right"><input type="submit" value="登陆" id="Submit1" name="Submit1"></td>
			</tr>
			</table>
			</td>
		<tr>
</form>
</font>
<% End IF %>
<% Else %>
	Response.Redirect("default.asp")
<% End IF %>

</body>
</html>
