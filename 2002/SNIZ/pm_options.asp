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
<!--#INCLUDE FILE="inc_top_short.asp" -->

<% Response.Buffer = True
if Request.Cookies(strUniqueID & "User")("Name") = "" Then %>
<meta http-equiv="Refresh" content="0; URL=<% =Request.ServerVariables("HTTP_REFERER") %>">
<center>
<br><br><br><br>
<font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><B>��½��̳����ʹ���������</b></font>
<br>
<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="<% =Request.ServerVariables("HTTP_REFERER") %>">������̳</font></a></p>
<br><br><br><br><br>
</center>
<% else 
if Request.QueryString("mode") = "setoptions" then
'## Forum_SQL
strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
strSql = strSql & " SET M_PMRECEIVE = " & Request.Form("statusstorage") & ", "
strSql = strSql & "     M_PMEMAIL = " & Request.Form("emailstorage")
strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & Request.Cookies(strUniqueID & "User")("Name") & "'"
strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_PASSWORD = '" & Request.Cookies(strUniqueID & "User")("PWord") & "'"

my_Conn.Execute(strSql)
if strSetCookieToForum = 1 then
   Response.Cookies("paging").Path = strCookieURL
end if
   Response.Cookies("paging")("outbox") = Request.Form("layoutstorage")
   Response.Cookies("paging").Expires = dateAdd("d", 360, strForumTimeAdjust)
	if request.cookies("paging")("outbox") = "" then request.cookies("paging")("outbox") = "single"

%>
<center>
<br>
<br>
<br>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>���Ļ������Ѹ���</b></font>
<br>

<br>
<br>
<br>
<br>
<br>

<% else 

		'## Forum_SQL
		strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.M_PASSWORD, " & strMemberTablePrefix & "MEMBERS.M_PMRECEIVE, " & strMemberTablePrefix & "MEMBERS.M_PMEMAIL "
		strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & Request.Cookies(strUniqueID & "User")("Name") & "'"
		strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_PASSWORD = '" & Request.Cookies(strUniqueID & "User")("PWord") & "'"

		set rs = my_Conn.Execute(strSql)

%>
<center>
<table width=100% border=0 cellspacing=1 cellpadding=4 bgcolor="<% =strTableBorderColor %>">
<form action="pm_options.asp?mode=setoptions" method="POST">
<TR bgcolor="<% =strHeadCellColor %>">
<TD><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>"><B>����/�ر� ���Ļ���ϢѶϢ</B></FONT></TD>
</tr>
<tr bgcolor="<% =strForumFirstCellColor %>">
  <td>
<font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
<b><% =strForumTitle %> ���Ļ���ϢĿǰ״̬Ϊ��<% if rs("M_PMRECEIVE") = "1" then %>����<% else %>�ر�<% end if %></b>��<br>����Դӵ���ѡ���<% if rs("M_PMRECEIVE") = "1" then %>�ر�<% else %>����<% end if %>���رպ��㽫���ղ������Ļ�ѶϢ<br>
<INPUT TYPE="RADIO" NAME="statusstorage" VALUE="1" <% if rs("M_PMRECEIVE") = "1" then Response.Write("checked")%>> �������Ļ�ѶϢ�� 
<INPUT TYPE="RADIO" NAME="statusstorage" VALUE="0" <% if rs("M_PMRECEIVE") = "0" then Response.Write("checked")%>> �ر����Ļ�ѶϢ��<br>
<font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�������ʱ�ص���ҳ�л����Ļ�ѶϢ��״̬</font>
</FONT>
</TD></TR>
<TR bgcolor="<% =strHeadCellColor %>">
  <TD><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>"><B>�����ʼ�֪ͨ</B></FONT>
    </TD>
</tr>
<tr bgcolor="<% =strForumFirstCellColor %>"><td><font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b><% =strForumTitle %> �����ڵ����յ����Ļ�ѶϢʱ������һ������ʼ�֪ͨ</b><br>
<INPUT TYPE="RADIO" NAME="emailstorage" VALUE="1" <% if rs("M_PMEMAIL") = "1" then Response.Write("checked")%>> ϣ���յ������ʼ�֪ͨ�� 
<INPUT TYPE="RADIO" NAME="emailstorage" VALUE="0" <% if rs("M_PMEMAIL") = "0" then Response.Write("checked")%>> ��Ҫ�յ�֪ͨѶϢ��</FONT>
</TD>
</TR>
<TR bgcolor="<% =strHeadCellColor %>">
<TD><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>"><B>�ռ���/������ �趨</B></FONT></TD>
</tr>
<tr bgcolor="<% =strForumFirstCellColor %>">
<td>
<font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
<b>Please select your layout preference.</b><br>
<INPUT TYPE="RADIO" NAME="layoutstorage" VALUE="single" <% if request.cookies("paging")("outbox") = "single" then Response.Write("checked")%>> ��ҳ���֣�&nbsp;<font size="1">�ռ���ͷ����������ͬһ��ҳ���</font><br>
<INPUT TYPE="RADIO" NAME="layoutstorage" VALUE="double" <% if request.cookies("paging")("outbox") = "double" then Response.Write("checked")%>> ˫ҳ���֣�&nbsp;&nbsp;<font size="1">�ռ���ͷ�����ֱ����������ҳ���</font><br>
<INPUT TYPE="RADIO" NAME="layoutstorage" VALUE="none" <% if request.cookies("paging")("outbox") = "none" then Response.Write("checked")%>> û���ռ��䣺&nbsp;<font size="1">û���ռ���.  ���͵���Ϣ�ǲ����洢�ġ�</font><br>
</FONT>
</Td></Tr>

</TABLE>
<P>
<INPUT TYPE="submit" VALUE="��������">
<% end if %>
<% end if %>
</P>
<!--#INCLUDE FILE="inc_footer_short.asp" -->