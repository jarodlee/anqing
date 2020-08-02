<!--#INCLUDE FILE="config.asp" --> 
<!--#INCLUDE FILE="inc_functions.asp" -->
<html>

<head>
<title><% =strForumTitle %></title>
</head>
<body bgColor="#FFFFFF" text="#000000">
<% 
set my_Conn = Server.CreateObject("ADODB.Connection") 
 my_Conn.Open strConnString 
	'## Forum_SQL - Get Origional Posting
	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strTablePrefix & "TOPICS.T_DATE, " & strTablePrefix & "TOPICS.T_SUBJECT, " & strTablePrefix & "TOPICS.T_AUTHOR, " & strTablePrefix & "TOPICS.TOPIC_ID, " & strTablePrefix & "TOPICS.T_MESSAGE "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS, " & strTablePrefix & "TOPICS "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "TOPICS.T_AUTHOR "
	strSql = strSql & " AND   " & strTablePrefix & "TOPICS.TOPIC_ID = " &  Request.QueryString("TOPIC_ID") 

	set rs4 = my_Conn.Execute (strSql)
	
	
	'## Forum_SQL - Get all topicsFrom DB
	strSql ="SELECT " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strTablePrefix & "REPLY.REPLY_ID, " & strTablePrefix & "REPLY.R_AUTHOR, " & strTablePrefix & "REPLY.TOPIC_ID, " & strTablePrefix & "REPLY.R_MESSAGE, " & strTablePrefix & "REPLY.R_DATE "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS, " & strTablePrefix & "REPLY "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "REPLY.R_AUTHOR "
	strSql = strSql & " AND   TOPIC_ID = " & Request.QueryString("TOPIC_ID") & " "
	strSql = strSql & " ORDER BY " & strTablePrefix & "REPLY.R_DATE"

	set rs3 = Server.CreateObject("ADODB.Recordset")
	rs3.cachesize = 20
	rs3.open  strSql, my_Conn, 3

		
		

%>

<Font Face="Times New Roman">
<p>
<font size="5"><b><% =rs4("T_Subject")%></font><br>主题作者：<% =rs4("M_NAME")%></b></p>

主题来自：<br>
<b><%=strForumTitle%></b><br>
<a href="<%=strForumURL%>/link.asp?TOPIC_ID=<%=Request.QueryString("TOPIC_ID") %>"><%=strForumURL%>/link.asp?TOPIC_ID=<%=Request.QueryString("TOPIC_ID") %></a>
<BR>

==========================================================
<BR><BR>
<b>标题：</b> <% =rs4("T_Subject")%><BR>
<b>日期：</b> <% =ChkDate(rs4("T_DATE")) %><BR>
<b>内容：</b><br><% =vbCrLf & formatStr(rs4("T_MESSAGE"))  %><BR>

<BR><BR><b><% if formatStr(rs3("R_MESSAGE")) = "" then %>没有回复
<% else %>
回复：
<% end if %></b>
<% i = 0 
	If rs3.EOF or rs3.BOF then  
		Response.Write ""
	Else
		rs3.movefirst
		do until rs3.EOF %>
		
<BR>----------------------------------------------------------


<BR><BR>		
<b>作者：</b> <% =rs3("M_NAME")%><BR>
<b>日期：</b> <% =ChkDate(rs3("R_DATE")) %><BR>
<b>内容：</b><br><% =vbCrLf & formatStr(rs3("R_MESSAGE")) %><BR>

<% 
rs3.MoveNext
		    strI = strI + 1
		    if strI = 2 then 
				strI = 0
			end if
		    rec = rec + 1
		loop
	end if
 %>
<BR>
==========================================================<BR>
￣讯息到此结束￣<BR><BR>

首页：<a href="<%=strHomeURL%>"><%=strHomeURL%></a><br>
<%=strForumTitle%>：<a href="<%=strForumURL%>"><%=strForumURL%></a>


</font>
<% my_conn.close %>
</body>
</html>
