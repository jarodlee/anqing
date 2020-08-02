<!--#INCLUDE FILE="config.asp" -->
<% If Session(strCookieURL & "Approval") = "15916941253" Then %>
<!--#INCLUDE file="inc_functions.asp" -->
<!--#INCLUDE file="inc_top.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<table border="0" width="100%">
  <tr>
    <td width="33%" align="left" nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    <img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="default.asp">返回论坛首页</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="admin_home.asp">论坛管理中心</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" height=15 width=15 border="0">&nbsp;<a href="admin_delete_topics.asp">会员贴子管理中心</a><br>
    </font></td>
  </tr>
</table>
<%
strDoDelete = "no"
	if request.querystring("m_id") <> "" then
    	strDoDelete = "yes"
    End if


Select Case strDoDelete
	Case "no"
 
    
strsql = "Select M_NAME, M_STATUS, MEMBER_ID, M_LEVEL, M_LAST_IP, M_POSTS from " & strTablePrefix & "MEMBERS"
set MemListRs = my_conn.execute(strsql)
	if MemListRs.eof then
    	response.write("No members Found")
    else
    %><font color="<%=strForumFontColor%>" size="<%=strDefaultFontSize%>" face="<%=strDefaultFontFace%>">
    <center><p>这功能将会批量删除某个会员的全部贴子(发表跟回复)，请在下面选择一个会员</p></centeR></font>
<table align="center" bgcolor="<%=strTableBorderColor%>" cellspacing=1 cellpadding=3>
	<tr>
    	<td bgcolor="<%=strHeadCellColor%>"><font color="<%=strHeadFontColor%>" face="<%=strDefaultFontFace%>" size="<%=strHeaderFontSize%>">会员名</font></td>
        <td bgcolor="<%=strHeadCellColor%>"><font color="<%=strHeadFontColor%>" face="<%=strDefaultFontFace%>" size="<%=strHeaderFontSize%>">登陆IP</font></td>
        <td bgcolor="<%=strHeadCellColor%>"><font color="<%=strHeadFontColor%>" face="<%=strDefaultFontFace%>" size="<%=strHeaderFontSize%>">会员等级</font></td>
        <td bgcolor="<%=strHeadCellColor%>"><font color="<%=strHeadFontColor%>" face="<%=strDefaultFontFace%>" size="<%=strHeaderFontSize%>">发表总数</font></td>
    </tr>
    <%
    		strDisplay = ""
            rec = 0
    	do until MemListRs.eof
        	if MemListRs("M_LEVEL") > 2 then
            	strMemberLevelA = "管理员"
            elseif MemListRs("M_LEVEL") > 1 and MemListRs("M_LEVEL") < 3 then
            	strMemberLevelA = "版主"
            else
            	strMemberLevelA = "一般会员"
            End if
            	if rec mod 2=0 then
                %>
        	<tr bgcolor="<%=strAltForumCellColor%>"><td>
            	<%else%>
            <tr bgcolor="<%=strForumCellColor%>"><td>
            	<%End if%>
            <font color="<%=strForumFontColor%>" size="<%=strDefaultFontSize%>" face="<%=strDefaultFontFace%>">
        	<a href="admin_delete_topics.asp?m_id=<%=MemListRs("MEMBER_ID")%>"><%=MemListRs("M_NAME")%></a></font></td><td>
            <font color="<%=strForumFontColor%>" size="<%=strDefaultFontSize%>" face="<%=strDefaultFontFace%>"><%=MemListRs("M_LAST_IP")%>
            </font></td><td>
            <font color="<%=strForumFontColor%>" size="<%=strDefaultFontSize%>" face="<%=strDefaultFontFace%>">
            <%=strMemberLevelA%>
            </font></td>
            <td><font color="<%=strForumFontColor%>" size="<%=strDefaultFontSize%>" face="<%=strDefaultFontFace%>"><%=MemListRS("M_POSTS")%></font></td>
            </tr>
            <%
            rec = rec + 1
        MemListRs.movenext
        Loop
    End if
set MemListRs = nothing
%>

</table>


<%
Case "yes"
	response.write("<p>")
	if request.querystring("delete") = "" then
%>
<font color="<%=strForumFontColor%>" size="<%=strDefaultFontSize%>" face="<%=strDefaultFontFace%>">
<center><p>是否要删除此会员的全部贴子?<br>请谨慎选择</p></center>

<center><a href="admin_delete_topics.asp?m_id=<%=request.querystring("m_id")%>&delete=yes">是</a> | <a href="admin_delete_topics.asp?m_id=<%=request.querystring("m_id")%>&delete=no">否</a></center>
</font>
<%
	elseif request.querystring("delete") = "yes" AND request.querystring("upcount") = "" then
    %>
    <font color="<%=strForumFontColor%>" size="<%=strDefaultFontSize%>" face="<%=strDefaultFontFace%>">
    <%
    	Call DoDeleteMemberTopics(request.querystring("m_id"))
		Call DoCounts()
%>

<center>已经成功删除此会员的全部贴子</center>
<center><a href="admin_delete_topics.asp">点击这里</a>返回会员贴子管理中心</center>
</font>
<%
    elseif request.querystring("delete") = "no" AND request.querystring("upcount") = "" then
%>
<font color="<%=strForumFontColor%>" size="<%=strDefaultFontSize%>" face="<%=strDefaultFontFace%>">
<center><p>会员贴子还没被删哦！</p><br><a href="admin_delete_topics.asp">点击这里</a>返回会员贴子管理中心</center>
</font>
<%
	elseif request.querystring("delete") = "yes" AND request.querystring("upcount") = "yes" then
    %>
    <font color="<%=strForumFontColor%>" size="<%=strDefaultFontSize%>" face="<%=strDefaultFontFace%>">
    <center><strong>已经成功删除此会员的全部贴子.</strong></center>
<center><a href="admin_delete_topics.asp">点击这里</a>返回会员贴子管理中心</center>
</font>
    <%
	end if
    response.write("<p>")
End Select

Sub DoDeleteMemberTopics(fMemId)
	strsql = "delete from " & strTablePrefix & "TOPICS where T_AUTHOR=" & fMemId
    My_Conn.execute(strsql)
    
    strsql = "delete from " & strTablePrefix & "REPLY where R_AUTHOR=" & fMemId
    My_Conn.execute(strsql)
    
    strsql = "update " & strTablePrefix & "MEMBERS set M_POSTS=0 WHERE MEMBER_ID=" & fMemId
    My_Conn.execute(strsql)
End Sub

Sub DoCounts()
response.write("<table align=center border=0>" &_
  "<tr>" & _
   " <td align=center colspan=2><p><b><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>更新主题统计...</font></b><br>" & _
    "&nbsp;</p></td>" & _
 " </tr>")
set rs = Server.CreateObject("ADODB.Recordset")
set rs1 = Server.CreateObject("ADODB.Recordset")

Response.Write "  <tr>" & vbCrLf
Response.Write "    <td align=right valign=top><font face='" &strDefaultFontFace & "'>主题数:</font></td>" & vbCrLf
Response.Write "    <td valign=top><font face='" &strDefaultFontFace & "'>"

'## Forum_SQL - Get contents of the Forum table related to counting
strSql = "SELECT FORUM_ID, F_TOPICS FROM " & strTablePrefix & "FORUM WHERE F_TYPE <> 1 "

rs.Open strSql, my_Conn
rs.MoveFirst
i = 0 

do until rs.EOF
i = i + 1

	'## Forum_SQL - count total number of topics in each forum in Topics table
	strSql = "SELECT count(FORUM_ID) AS cnt "
	strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
	strSql = strSql & " WHERE FORUM_ID = " & rs("FORUM_ID")

	rs1.Open strSql, my_Conn

	if rs1.EOF or rs1.BOF then
		intF_TOPICS = 0
	Else
		intF_TOPICS = rs1("cnt")
	End if
	
	strSql = "UPDATE " & strTablePrefix & "FORUM "
	strSql = strSql & " SET F_TOPICS = " & intF_TOPICS
	strSql = strSql & " WHERE FORUM_ID = " & rs("FORUM_ID")
	
	my_conn.execute(strSql)
	
	rs1.Close
	rs.MoveNext
	Response.Write "."
	if i = 80 then 
		Response.Write "    <br>" & vbCrLf
		i = 0
	End if
loop
rs.Close

Response.Write "    </font></td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "  <tr>" & vbCrLf
Response.Write "    <td align=right valign=top><font face='" &strDefaultFontFace & "'>回复数:</font></td>" & vbCrLf
Response.Write "    <td valign=top><font face='" & strDefaultFontFace & "'>"

'## Forum_SQL
strSql = "SELECT TOPIC_ID, T_REPLIES FROM " & strTablePrefix & "TOPICS"

rs.Open strSql, my_Conn
i = 0 

do until rs.EOF
i = i + 1

	'## Forum_SQL - count total number of replies in Topics table
	strSql = "SELECT count(REPLY_ID) AS cnt "
	strSql = strSql & " FROM " & strTablePrefix & "REPLY "
	strSql = strSql & " WHERE TOPIC_ID = " & rs("TOPIC_ID")

	rs1.Open strSql, my_Conn
	if rs1.EOF or rs1.BOF or (rs1("cnt") = 0) then
		intT_REPLIES = 0
		
		set rs2 = Server.CreateObject("ADODB.Recordset")

		'## Forum_SQL - Get post_date and author from Topic
		strSql = "SELECT T_AUTHOR, T_DATE "
		strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
		strSql = strSql & " WHERE TOPIC_ID = " & rs("TOPIC_ID") & " "
				
		set rs2 = my_Conn.Execute (strSql)
			
		if not(rs2.eof or rs2.bof) then
			strLast_Post = rs2("T_DATE")
			strLast_Post_Author = rs2("T_AUTHOR")
		else
			strLast_Post = ""
			strLast_Post_Author = ""
		end if
				
		rs2.Close
		set rs2 = nothing
		
	Else
		intT_REPLIES = rs1("cnt")
		
		'## Forum_SQL - Get last_post and last_post_author for Topic
		strSql = "SELECT R_DATE, R_AUTHOR "
		strSql = strSql & " FROM " & strTablePrefix & "REPLY "
		strSql = strSql & " WHERE TOPIC_ID = " & rs("TOPIC_ID") & " "
		strSql = strSql & " ORDER BY R_DATE DESC"
				
		set rs3 = my_Conn.Execute (strSql)
			
		if not(rs3.eof or rs3.bof) then
			rs3.movefirst
			strLast_Post = rs3("R_DATE")
			strLast_Post_Author = rs3("R_AUTHOR")
		else
			strLast_Post = ""
			strLast_Post_Author = ""
		end if
	
		rs3.close
		set rs3 = nothing
		
	End if
	
	strSql = "UPDATE " & strTablePrefix & "TOPICS "
	strSql = strSql & " SET T_REPLIES = " & intT_REPLIES
	if strLast_Post <> "" then 
		strSql = strSql & ", T_LAST_POST = '" & strLast_Post & "'"
		if strLast_Post_Author <> "" then 

			strSql = strSql & ", T_LAST_POST_AUTHOR = " & strLast_Post_Author 

		end if
	end if
	strSql = strSql & " WHERE TOPIC_ID = " & rs("TOPIC_ID")
	
	my_conn.execute(strSql)
		
	rs1.Close
	rs.MoveNext
	Response.Write "."
	if i = 80 then 
		Response.Write "    <br>" & vbCrLf
		i = 0
	End if
loop
rs.Close

Response.Write "    </font></td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "  <tr>" & vbCrLf
Response.Write "    <td align=right valign=top><font face='" & strDefaultFontFace & "'>论坛回复:</font></td>" & vbCrLf
Response.Write "    <td valign=top><font face='" &strDefaultFontFace & "'>"

'## Forum_SQL - Get values from Forum table needed to count replies
strSql = "SELECT FORUM_ID, F_COUNT FROM " & strTablePrefix & "FORUM WHERE F_TYPE <> 1 "

rs.Open strSql, my_Conn, 2, 2

do until rs.EOF

	'## Forum_SQL - Count total number of Replies
	strSql = "SELECT Sum(" & strTablePrefix & "TOPICS.T_REPLIES) AS SumOfT_REPLIES, Count(" & strTablePrefix & "TOPICS.T_REPLIES) AS cnt "
	strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
	strSql = strSql & " WHERE " & strTablePrefix & "TOPICS.FORUM_ID = " & rs("FORUM_ID")

	rs1.Open strSql, my_Conn
	
	if rs1.EOF or rs1.BOF then
		intF_COUNT = 0
		intF_TOPICS = 0
	Else
		intF_COUNT = rs1("cnt") + rs1("SumOfT_REPLIES")
		intF_TOPICS = rs1("cnt") 
	End if
	If IsNull(intF_COUNT) then intF_COUNT = 0 
	if IsNull(intF_TOPICS) then intF_TOPICS = 0 
	
	'## Forum_SQL - Get last_post and last_post_author for Forum
	strSql = "SELECT T_LAST_POST, T_LAST_POST_AUTHOR "
	strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
	strSql = strSql & " WHERE FORUM_ID = " & rs("FORUM_ID") & " "
	strSql = strSql & " ORDER BY T_LAST_POST DESC"

	set rs2 = my_Conn.Execute (strSql)
			
	if not (rs2.eof or rs2.bof) then
		strLast_Post = rs2("T_LAST_POST")
		strLast_Post_Author = rs2("T_LAST_POST_AUTHOR")
	else
		strLast_Post = ""
		strLast_Post_Author = ""
	end if
			
	rs2.Close
	set rs2 = nothing
	
	strSql = "UPDATE " & strTablePrefix & "FORUM "
	strSql = strSql & " SET F_COUNT = " & intF_COUNT
	strSql = strSql & ",  F_TOPICS = " & intF_TOPICS
	if strLast_Post <> "" then 
		strSql = strSql & ", F_LAST_POST = '" & strLast_Post & "' "
		if strLast_Post_Author <> "" then 



			strSql = strSql & ", F_LAST_POST_AUTHOR = " & strLast_Post_Author



		end if
	end if
	strSql = strSql & " WHERE FORUM_ID = " & rs("FORUM_ID")
	
	'Response.Write strSql
	'Response.End
	
	my_conn.execute(strSql)
		
	rs1.Close
	rs.MoveNext
	Response.Write "."
	if i = 80 then 
		Response.Write "    <br>" & vbCrLf
		i = 0
	End if	
loop
rs.Close

Response.Write "    </font></td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "  <tr>" & vbCrLf
Response.Write "    <td align=right valign=top><font face='" &strDefaultFontFace & "'>统计:</font></td>" & vbCrLf
Response.Write "    <td valign=top><font face='" &strDefaultFontFace & "'>"

'## Forum_SQL - Total of Topics
strSql = "SELECT Sum(" & strTablePrefix & "FORUM.F_TOPICS) "
strSql = strSql & " AS SumOfF_TOPICS "
strSql = strSql & " FROM " & strTablePrefix & "FORUM WHERE F_TYPE <> 1 "

rs.Open strSql, my_Conn

Response.Write "主题数: " & RS("SumOfF_TOPICS") & "<br>" & vbCrLf

'## Forum_SQL - Write total Topics to Totals table
strSql = "UPDATE " & strTablePrefix & "TOTALS "
strSql = strSql & " SET T_COUNT = " & rs("SumOfF_TOPICS")

rs.Close

my_Conn.Execute strSql

'## Forum_SQL - Total all the replies for each topic
strSql = "SELECT Sum(" & strTablePrefix & "FORUM.F_COUNT) "
strSql = strSql & " AS SumOfF_COUNT "
strSql = strSql & " FROM " & strTablePrefix & "FORUM WHERE F_TYPE <> 1 "

set rs = my_Conn.Execute (strSql)
'rs.Open strSql, my_Conn

if rs("SumOfF_COUNT") <> "" then
	Response.Write "发表总数: " & RS("SumOfF_COUNT") & "<br>" & vbCrLf
	strSumOfF_COUNT = rs("SumOfF_COUNT")
else
	Response.Write "发表总数: 0<br>" & vbCrLf
	strSumOfF_COUNT = "0"
end if

'## Forum_SQL - Write total replies to the Totals table
strSql = "UPDATE " & strTablePrefix & "TOTALS "
strSql = strSql & " SET P_COUNT = " & strSumOfF_COUNT

rs.Close

my_Conn.Execute strSql

'## Forum_SQL - Total number of users
strSql = "SELECT Count(MEMBER_ID) "
strSql = strSql & " AS CountOf "
strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS"

rs.Open strSql, my_Conn

Response.Write "会员总数: " & RS("Countof") & "<br>" & vbCrLf

'## Forum_SQL - Write total number of users to Totals table
strSql = " UPDATE " & strTablePrefix & "TOTALS "
strSql = strSql & " SET U_COUNT = " & cint(RS("Countof"))

rs.Close

my_Conn.Execute strSql

Response.Write "    </font></td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "  <tr>" & vbCrLf
Response.Write "    <td align=center colspan=2>&nbsp;<br>" & vbCrLf
Response.Write "    <b><font face='" & strDefaultFontFace & "' size=" & strHeaderFontSize & ">主题统计更新完毕！</font></b></font></td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf

'on error resume next

set rs = nothing
set rs1 = nothing
response.write("</table>")
End Sub
%>
<!--#INCLUDE file="inc_footer.asp" -->
<% Else %>
<% Response.Redirect "admin_login.asp" %>
<% End IF %>
