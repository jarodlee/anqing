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
<% server.scripttimeout = 6000 %>
<!--#INCLUDE FILE="config.asp" -->
<% if Session(strCookieURL & "Approval") = "15916941253" then %>
<!--#INCLUDE FILE="inc_functions.asp" -->
<!--#INCLUDE file="inc_top.asp" -->
<% 
My_ID = request.querystring("id")
My_Mode = request.querystring("mode")
if My_Mode = "" then
%>
<table border="0" width="100%">
  <tr>
 <td width="33%" align="left" nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
 <img src="images/icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="default.asp">������̳��ҳ</a>&nbsp;<img src="images/icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="admin_home.asp">��̳��������</a>&nbsp;<img src="images/icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="admin_emaillist.asp">��Ա�����ʼ��б�</a>&nbsp;<img src="images/icon_folder_open_topic.gif" height=15 width=15 border="0">&nbsp;��Ա�ʼ�����<br>
 </font></td>
  </tr>
</table>
<BR>
<% 
strSql = "SELECT * FROM " &strMemberTablePrefix & "SPAM ORDER BY ARCHIVE ASC"
set rs = Server.CreateObject("ADODB.Recordset")
rs.open  strSql, My_Conn, 3
%>
<TABLE width="95%" border="0" cellspacing="1" cellpadding="3" align="center" bgcolor="<% =strPopUpBorderColor %>">
<TR ALIGN="CENTER">
<TD BGCOLOR=<% =strHeadCellColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">״̬</FONT></TD>
<TD BGCOLOR=<% =strHeadCellColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">ѶϢ����</FONT></TD>
<TD BGCOLOR=<% =strHeadCellColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">����</FONT></TD>
<TD BGCOLOR=<% =strHeadCellColor %>><a href="admin_emailmanager.asp?mode=compose"><img src="image/add-g2.gif" alt="������ѶϢ" border="0" hspace="0"></a></TD>
</TR>
<%
On Error Resume Next
RS.MoveFirst
do while Not RS.eof                       
ARCHIVED = rs("ARCHIVE")
if ARCHIVED = "1" then
ARCHIVED = "����"
else
ARCHIVED = "��ʱ"
end if
if rs("F_SENT") <> "" then
F_SENT = ChkDate(rs("F_SENT"))
else
F_SENT = "-" 
end if
 %>
<TR VALIGN=TOP>
<td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><%= ARCHIVED %></FONT></TD>
<td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><input type="hidden" name="ID" value="<%=RS("ID")%>"><a href="admin_emailmanager.asp?mode=edit&id=<%=RS("id")%>"><%=RS("SUBJECT")%></a></FONT></TD>
<td ALIGN="CENTER" bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% =F_SENT %></FONT></TD>
<td bgcolor="<% =strForumCellColor %>" align="right"> <a href="admin_emailmanager.asp?mode=edit&ID=<% =rs("ID") %>"><img src="image/pen.gif" alt="�༭ѶϢ" border="0" hspace="0"></a>
  <a href="admin_emailmanager.asp?mode=update&ID=<% =rs("ID") %>&ARCHIVE=2"><img src="image/del2.gif" alt="ɾ��ѶϢ" border="0" hspace="0"></a></td>
</TR>
<%
RS.MoveNext
loop%>
</table>
<%
set rs = nothing
%>
<BR><BR><!--#INCLUDE file="inc_footer.asp" -->

<% elseif My_Mode = "update" then '---------------------------------update  %>
<%
if request.querystring("mode") = "update" then
%>
<%  '--------------------------------------------------delete
if request.querystring("ARCHIVE")= "2" then
 
		set conn = server.createobject("adodb.connection")
	      	conn.Open My_Conn
		For each record in request("ID")
    		sqlstmt = "DELETE * from FORUM_SPAM WHERE ID=" & My_ID
			Set kRS = conn.execute(sqlstmt)
		Next
 
	set kRS = nothing
'----------------------------------------------------/delete
%>
<BR><BR>

<font face="<% =strDefaultFontFace %>"><h1>ѶϢ��ɾ��</h1><BR><BR>
ɾ����ɣ�<a href="admin_emailmanager.asp">�ص�ѶϢ�б�</a>��</font>

<BR><BR>
<!--#INCLUDE file="inc_footer.asp" -->
	<%
	response.end
end if
%>
<% 
My_ID = request.querystring("ID") 
	strSQL3="select * from FORUM_SPAM where id=" & My_ID
	set kRS=Server.CreateObject("ADODB.Recordset")
	kRS.Open strSQL3, My_Conn, 1, 3
kRS("SUBJECT") = request.querystring("SUBJECT")
kRS("MESSAGE") = request.querystring("MESSAGE")	
kRS("ARCHIVE") = request.querystring("ARCHIVE")

	kRS.Update
%>
<BR><BR><font face="<% =strDefaultFontFace %>"><h1>ѶϢ�Ѹ���</h1><BR><BR>
������ɣ�<a href="admin_emailmanager.asp">�ص�ѶϢ�б�</a>��</font>
<%
set kRS = nothing
%>

<BR><BR><!--#INCLUDE file="inc_footer.asp" -->
<% end if  %>
<% '----------------------------------/update
elseif My_Mode = "compose" then 
'--------------------------------------------/compose
%>
<h2>�����ѶϢ</h2>
<form action="admin_emailmanager.asp">
<input type="hidden" name="mode" value="save">
<table bordercolor="<% =strTableBorderColor %>" border="0" cellspacing="1" cellpadding="5">
<tr>
<td><font face="<% =strDefaultFontFace %>">���⣺</font></td><td><input type="text" name="SUBJECT" size="50"></td>
</tr>
<tr>
<td colspan="2"><font face="<% =strDefaultFontFace %>">�ż����ݣ�</font></td>
</tr>
<tr>
<td colspan="2" align="center"><textarea name="MESSAGE" cols="50" rows="10" wrap="PHYSICAL"></textarea></td>
</tr>
</table>

<font face="<% =strDefaultFontFace %>">�Ժ��ַ�ʽ�����ѶϢ��</font>
<font face="����, Arial, Helvetica" size="2"> 
 <select name="ARCHIVE" size="1">
  <option value="0" SELECTED>&nbsp;��ʱ�б�</option>
  <option value="1">&nbsp;����</option>
</select>
 </font>

 &nbsp;<input type="Submit" value="����">&nbsp;<input type="reset">
 
<% '--------------------------------------------/compose
elseif My_Mode = "save" then
'---------------------------------save  %>
<%
strSubject = replace(request.querystring("SUBJECT"),"'","''")
strMessage = replace(request.querystring("MESSAGE"),"'","''")
strArchive = request.querystring("ARCHIVE")

	set conn = server.createobject ("adodb.connection")
	conn.open My_Conn
	conn.Execute "insert into FORUM_SPAM (SUBJECT, MESSAGE, F_SENT, ARCHIVE) values (" _
		& "'" & strSubject & "', " _
		& "'" & strMessage & "', " _ 
		& "'" & DateToStr(now()) & "', " _		
		& "'" & strArchive & "')"
%>
<BR><BR><font face="<% =strDefaultFontFace %>"><h2>ѶϢ�Ѵ���</h2><BR><BR>
����ɹ���<a href="admin_emailmanager.asp">�ص���Ա�ʼ�����</a>��</font>		
<BR><BR><!--#INCLUDE file="inc_footer.asp" -->

<% '--------------------------------------------/save
elseif My_Mode = "edit" then
'---------------------------------edit  %>
<%
strSql2 = "SELECT * FROM " &strMemberTablePrefix & "SPAM WHERE ID =" & My_ID
set rsSP = Server.CreateObject("ADODB.Recordset")
rsSP.open  strSql2, My_Conn, 3
mySUBJECT = Server.HTMLEncode(rsSP("SUBJECT"))
myMESSAGE = rsSP("MESSAGE")
%>
<h2>�༭ѶϢ</h2>
<form action="admin_emailmanager.asp"><input type="hidden" name="mode" value="update"><input type="hidden" name="ID" value="<%= rsSP("ID") %>">
<table bordercolor="<% =strTableBorderColor %>" border="1" cellspacing="0" cellpadding="5">
<tr>
<td><font face="arial">���⣺</font></td><td><input type="text" name="SUBJECT" size="50" value="<%= mySUBJECT%>"></td>
</tr>
<tr>
<td colspan="2"><font face="arial">�ż����ݣ�</font></td>
</tr>
<tr>
<td colspan="2" align="center"><textarea name="MESSAGE" cols="50" rows="10" wrap="PHYSICAL"><%= myMESSAGE %></textarea></td>
</tr>
</table>
<font face="arial">ѶϢ״̬��</font>&nbsp; 
<font face="����, Arial, Helvetica" size="2"> 
 <select name="ARCHIVE" size="1">
  <option value="0" SELECTED>&nbsp;��ʱ�б�</option>
  <option value="1">&nbsp;����</option>
  <option value="2">&nbsp;ɾ��</option>
</select>
 </font>
 &nbsp;<input type="Submit" value="�༭">&nbsp;<input type="reset">
<%
set rsSP = nothing
%>
<BR><BR><!--#INCLUDE file="inc_footer.asp" -->
<% end if '--------------------end edit %>
<% else %>
<%Response.Redirect "admin_login.asp" %>
<% end iF %>
