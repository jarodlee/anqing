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
%>
<!--#INCLUDE FILE="config.asp" -->
<!--#INCLUDE FILE="inc_functions.asp" -->
<!--#INCLUDE file="inc_top.asp" -->
<table border="0" width="100%">
  <tr>
    <td width="33%" align="left" nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    <img src="<%=strImageURL %>icon_folder_open.gif" alt="������̳��ҳ" height=15 width=15 border="0">&nbsp;<a href="default.asp">������̳��ҳ</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" alt="����Ա��½" height=15 width=15 border="0">&nbsp;����Ա��½
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
<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">��½�ɹ���</font></p>
<% Session(strCookieURL & "Approval") = "15916941253" %>
<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="admin_home.asp" target="_top">�����������̳</a></font></p>
<meta http-equiv="Refresh" content="0; URL=admin_home.asp">
<!--#INCLUDE file="inc_footer.asp" -->
<% Response.End %>
<% Else %>
<center>

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">�������⣡</font></p>

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">����Ȩ����</font></p>

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�������Ϊ���ѶϢ�Ǵ���ģ������³���һ��</font></p>

</center>
<% End IF %>
<% Else %>
	&nbsp;
<% End IF %>
<form action="admin_login.asp" method="post" id=Form1 name=Form1>
<table border="0" cellspacing="1" cellpadding="4" align="center" bgcolor="<% =strTableBorderColor %>">
  <tr>
    <td align="center" colspan="2" bgcolor="<% =strHeadCellColor %>"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">����Ա��½</font></b></td>
  </tr>
  <tr>
	<td align="right" nowrap bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">������֣�</font></td>
	<td bgcolor="<% =strForumCellColor %>"><input type="text" name="Name"></td>
  </tr>
  <tr>
	<td align="right" nowrap nowrap bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">������룺</font></td>
	<td bgcolor="<% =strForumCellColor %>"><input type="Password" name="Password"></td>
  </tr>
  <tr>
    <td colspan="2" align="right" bgcolor="<% =strForumCellColor %>"><input type="submit" value="��½" id="Submit1" name="Submit1"></td>
  </tr>
</table>
</form>

</font>
<!--#INCLUDE file="inc_footer.asp" -->
