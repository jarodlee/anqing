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
 <img src="images/icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="default.asp">返回论坛首页</a>&nbsp;<img src="images/icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="admin_home.asp">论坛管理中心</a>&nbsp;<img src="images/icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="admin_emaillist.asp">会员电子邮件列表</a>&nbsp;<img src="images/icon_folder_open_topic.gif" height=15 width=15 border="0">&nbsp;会员邮件管理<br>
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
<TD BGCOLOR=<% =strHeadCellColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">状态</FONT></TD>
<TD BGCOLOR=<% =strHeadCellColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">讯息标题</FONT></TD>
<TD BGCOLOR=<% =strHeadCellColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">构成</FONT></TD>
<TD BGCOLOR=<% =strHeadCellColor %>><a href="admin_emailmanager.asp?mode=compose"><img src="image/add-g2.gif" alt="加入新讯息" border="0" hspace="0"></a></TD>
</TR>
<%
On Error Resume Next
RS.MoveFirst
do while Not RS.eof                       
ARCHIVED = rs("ARCHIVE")
if ARCHIVED = "1" then
ARCHIVED = "档案"
else
ARCHIVED = "即时"
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
<td bgcolor="<% =strForumCellColor %>" align="right"> <a href="admin_emailmanager.asp?mode=edit&ID=<% =rs("ID") %>"><img src="image/pen.gif" alt="编辑讯息" border="0" hspace="0"></a>
  <a href="admin_emailmanager.asp?mode=update&ID=<% =rs("ID") %>&ARCHIVE=2"><img src="image/del2.gif" alt="删除讯息" border="0" hspace="0"></a></td>
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

<font face="<% =strDefaultFontFace %>"><h1>讯息已删除</h1><BR><BR>
删除完成！<a href="admin_emailmanager.asp">回到讯息列表</a>。</font>

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
<BR><BR><font face="<% =strDefaultFontFace %>"><h1>讯息已更新</h1><BR><BR>
更新完成！<a href="admin_emailmanager.asp">回到讯息列表</a>。</font>
<%
set kRS = nothing
%>

<BR><BR><!--#INCLUDE file="inc_footer.asp" -->
<% end if  %>
<% '----------------------------------/update
elseif My_Mode = "compose" then 
'--------------------------------------------/compose
%>
<h2>添加新讯息</h2>
<form action="admin_emailmanager.asp">
<input type="hidden" name="mode" value="save">
<table bordercolor="<% =strTableBorderColor %>" border="0" cellspacing="1" cellpadding="5">
<tr>
<td><font face="<% =strDefaultFontFace %>">标题：</font></td><td><input type="text" name="SUBJECT" size="50"></td>
</tr>
<tr>
<td colspan="2"><font face="<% =strDefaultFontFace %>">信件内容：</font></td>
</tr>
<tr>
<td colspan="2" align="center"><textarea name="MESSAGE" cols="50" rows="10" wrap="PHYSICAL"></textarea></td>
</tr>
</table>

<font face="<% =strDefaultFontFace %>">以何种方式储存此讯息？</font>
<font face="宋体, Arial, Helvetica" size="2"> 
 <select name="ARCHIVE" size="1">
  <option value="0" SELECTED>&nbsp;即时列表</option>
  <option value="1">&nbsp;档案</option>
</select>
 </font>

 &nbsp;<input type="Submit" value="储存">&nbsp;<input type="reset">
 
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
<BR><BR><font face="<% =strDefaultFontFace %>"><h2>讯息已储存</h2><BR><BR>
储存成功！<a href="admin_emailmanager.asp">回到会员邮件管理</a>。</font>		
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
<h2>编辑讯息</h2>
<form action="admin_emailmanager.asp"><input type="hidden" name="mode" value="update"><input type="hidden" name="ID" value="<%= rsSP("ID") %>">
<table bordercolor="<% =strTableBorderColor %>" border="1" cellspacing="0" cellpadding="5">
<tr>
<td><font face="arial">标题：</font></td><td><input type="text" name="SUBJECT" size="50" value="<%= mySUBJECT%>"></td>
</tr>
<tr>
<td colspan="2"><font face="arial">信件内容：</font></td>
</tr>
<tr>
<td colspan="2" align="center"><textarea name="MESSAGE" cols="50" rows="10" wrap="PHYSICAL"><%= myMESSAGE %></textarea></td>
</tr>
</table>
<font face="arial">讯息状态：</font>&nbsp; 
<font face="宋体, Arial, Helvetica" size="2"> 
 <select name="ARCHIVE" size="1">
  <option value="0" SELECTED>&nbsp;即时列表</option>
  <option value="1">&nbsp;档案</option>
  <option value="2">&nbsp;删除</option>
</select>
 </font>
 &nbsp;<input type="Submit" value="编辑">&nbsp;<input type="reset">
<%
set rsSP = nothing
%>
<BR><BR><!--#INCLUDE file="inc_footer.asp" -->
<% end if '--------------------end edit %>
<% else %>
<%Response.Redirect "admin_login.asp" %>
<% end iF %>
