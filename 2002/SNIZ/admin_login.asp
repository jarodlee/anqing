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
<!--#INCLUDE file="inc_top.asp" -->
<table border="0" width="100%">
  <tr>
    <td width="33%" align="left" nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    <img src="<%=strImageURL %>icon_folder_open.gif" alt="返回论坛首页" height=15 width=15 border="0">&nbsp;<a href="default.asp">返回论坛首页</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" alt="管理员登陆" height=15 width=15 border="0">&nbsp;管理员登陆
    </font></td>
  </tr>
</table>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
<%
Name = Request.Form("Name")
Password = Request.Form("Password")

RequestMethod = Request.ServerVariables("Request_method")

IF RequestMethod = "POST" Then

	'## Forum_SQL
	strSql = "SELECT MEMBER_ID "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE M_NAME = '" & Name & "' AND "
	strSql = strSql & "       M_PASSWORD = '" & Password & "' AND "
	strSql = strSql & "       M_LEVEL = 3"
	
	Set dbRs = my_Conn.Execute(strSql)
		
	If not(dbRS.EOF) and ChkQuoteOk(Name) and ChkQuoteOk(Password) Then 
%>
<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">登陆成功！</font></p>
<% Session(strCookieURL & "Approval") = "15916941253" %>
<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="admin_home.asp" target="_top">按这里进入论坛</a></font></p>
<meta http-equiv="Refresh" content="0; URL=admin_home.asp">
<!--#INCLUDE file="inc_footer.asp" -->
<% Response.End %>
<% Else %>
<center>

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">发生问题！</font></p>

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">你无权进入</font></p>

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">如果你认为这个讯息是错误的，请重新尝试一遍</font></p>

</center>
<% End IF %>
<% Else %>
	&nbsp;
<% End IF %>
<form action="admin_login.asp" method="post" id=Form1 name=Form1>
<table border="0" cellspacing="1" cellpadding="4" align="center" bgcolor="<% =strTableBorderColor %>">
  <tr>
    <td align="center" colspan="2" bgcolor="<% =strHeadCellColor %>"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">管理员登陆</font></b></td>
  </tr>
  <tr>
	<td align="right" nowrap bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">你的名字：</font></td>
	<td bgcolor="<% =strForumCellColor %>"><input type="text" name="Name"></td>
  </tr>
  <tr>
	<td align="right" nowrap nowrap bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">你的密码：</font></td>
	<td bgcolor="<% =strForumCellColor %>"><input type="Password" name="Password"></td>
  </tr>
  <tr>
    <td colspan="2" align="right" bgcolor="<% =strForumCellColor %>"><input type="submit" value="登陆" id="Submit1" name="Submit1"></td>
  </tr>
</table>
</form>

</font>
<!--#INCLUDE file="inc_footer.asp" -->
