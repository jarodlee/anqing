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
<% If Session(strCookieURL & "Approval") = "15916941253" Then %> 
<!--#INCLUDE FILE="inc_functions.asp" -->
<!--#INCLUDE file="inc_top.asp" -->
<table border="0" width="100%">
  <tr>
    <td width="33%" align="left" nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    <img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="default.asp">������̳��ҳ</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="admin_home.asp">��̳��������</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" height=15 width=15 border="0">&nbsp;������̳״̬ͳ��<br>
    </font></td>
  </tr>
</table>

<table align=center border=0>
  <tr>
    <td align=center colspan=2><p><b><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">���ڸ�����̳״̬.....</font></b><br>
    &nbsp;</p></td>
  </tr>
<%
'Server.ScriptTimeout = 6000

set rs = Server.CreateObject("ADODB.Recordset")
set rs1 = Server.CreateObject("ADODB.Recordset")

Response.Write "  <tr>" & vbCrLf
Response.Write "    <td align=right valign=top><font face='" &strDefaultFontFace & "'>������:</font></td>" & vbCrLf
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
Response.Write "    <td align=right valign=top><font face='" &strDefaultFontFace & "'>��������:</font></td>" & vbCrLf
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
Response.Write "    <td align=right valign=top><font face='" & strDefaultFontFace & "'>��̳����:</font></td>" & vbCrLf
Response.Write "    <td valign=top><font face='" &strDefaultFontFace & "'>"

'## Forum_SQL - Get values from Forum table needed to count replies
strSql = "SELECT FORUM_ID, F_COUNT FROM " & strTablePrefix & "FORUM WHERE F_TYPE <> 1 "

rs.Open strSql, my_Conn, 2, 2

do until rs.EOF
on error resume next
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
	
	my_conn.execute strSql,,adCmdText + adexecuteNoRecords
		
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
Response.Write "    <td align=right valign=top><font face='" &strDefaultFontFace & "'>ͳ��:</font></td>" & vbCrLf
Response.Write "    <td valign=top><font face='" &strDefaultFontFace & "'>"

'## Forum_SQL - Total of Topics
strSql = "SELECT Sum(" & strTablePrefix & "FORUM.F_TOPICS) "
strSql = strSql & " AS SumOfF_TOPICS "
strSql = strSql & " FROM " & strTablePrefix & "FORUM WHERE F_TYPE <> 1 "

rs.Open strSql, my_Conn

Response.Write "�ظ�: " & RS("SumOfF_TOPICS") & "<br>" & vbCrLf

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
	Response.Write "����: " & RS("SumOfF_COUNT") & "<br>" & vbCrLf
	strSumOfF_COUNT = rs("SumOfF_COUNT")
else
	Response.Write "Total Posts: 0<br>" & vbCrLf
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

Response.Write "��Ա����: " & RS("Countof") & "<br>" & vbCrLf

'## Forum_SQL - Write total number of users to Totals table
strSql = " UPDATE " & strTablePrefix & "TOTALS "
strSql = strSql & " SET U_COUNT = " & cint(RS("Countof"))

rs.Close

my_Conn.Execute strSql

Response.Write "    </font></td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf
Response.Write "  <tr>" & vbCrLf
Response.Write "    <td align=center colspan=2>&nbsp;<br>" & vbCrLf
Response.Write "    <b><font face='" & strDefaultFontFace & "' size=" & strHeaderFontSize & ">�������</font></b></font></td>" & vbCrLf
Response.Write "  </tr>" & vbCrLf

'on error resume next

set rs = nothing
set rs1 = nothing
%>
</table>
&nbsp;<meta http-equiv="Refresh" content="60; URL=admin_home.asp">

<!--#INCLUDE file="inc_footer.asp" -->
<% Else %>
<% Response.Redirect "admin_login.asp" %>
<% End IF %>
