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

blnSetup = Request.Form("setup")
%>
<!--#INCLUDE FILE="config.asp" -->
<!--#INCLUDE FILE="inc_functions.asp" -->
<html>

<head>
<title>Forum-Setup Page</title>
<meta name="copyright" content="��Դ����վ(www.99ss.net)">
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
<p align="center"><font face="����, Arial, Helvetica" size="4">��½�ɹ���</font></p>
<% Session(strCookieURL & "Approval") = "15916941253" %>
<p>&nbsp;</p>
<p align="center"><p align="center"><font face="����, Arial, Helvetica" size="2"><a href="setup.asp?RC=3" target="_top">Click here to Continue.</a></font></p>
<meta http-equiv="Refresh" content="0; URL=setup.asp?RC=3">
<% Response.End %>
<% else %>
<div align=center><center>
	<p><font face="����, Arial, Helvetica" size="4">There has been a problem !</font></p>
</center></div>

<form action="setup_login.asp" method="post" id=Form1 name=Form1>
<input type="hidden" name="setup" value="Y">
<table width="50%" height="50%" align="center" border="0" cellspacing="0" cellpadding="5">
		<tr>
			<td bgColor=navyblue align="center"><p align="center"><font face="����, Arial, Helvetica" size="2"><b>����Ȩ����</b></font></p></td>
		</tr>
		<tr>
			<td bgColor=navyblue align="left"><p><font face="����, Arial, Helvetica" size="2">I�������Ϊ���ѶϢ�Ǵ���ģ������³���һ��</font></p></td>
		</tr>
		<tr>
			<td>
			<table border="0" cellspacing="0" cellpadding="5" border="0" cellspacing="2" cellpadding="0" align="center">
			<tr>
			  <td align="center" colspan="2" bgColor=navyblue><b><font face="����, Arial, Helvetica" size="2">����Ա��½</font></b></td>
			</tr>
			<tr>
			  <td align="right" nowrap><b><font face="����, Arial, Helvetica" size="2">���ƣ�</font></b></td>
			  <td><input type="text" name="Name"></td>
			</tr>
			<tr>
			  <td align="right" nowrap nowrap><b><b><font face="����, Arial, Helvetica" size="2">���룺</font></b></td>
			  <td><input type="Password" name="Password"></td>
			</tr>
			<tr>
			  <td colspan="2" align="right"><input type="submit" value="��½" id="Submit1" name="Submit1"></td>
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
