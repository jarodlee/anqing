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
<!--#INCLUDE FILE="inc_top.asp" -->
<% 
if strAuthType = "db" then
	strDBNTUserName = Request.Form("UserName")
end if

set rs = Server.CreateObject("ADODB.RecordSet")

err_Msg = ""
ok = "" 

if Request.Form("Method_Type") = "Edit" then
	member = cint(ChkUser(strDBNTUserName, Request.Form("Password")))
	Select Case Member 
		case 0 '## Invalid Pword
			Go_Result "错误的用户名跟密码", 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
		case 1 '## Author of Post so OK
			'## Do Nothing
		case 2 '## Normal User - Not Authorised
			Go_Result "只有管理员、版主或作者才能修改此文章", 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
		case 3 '## Moderator so OK - check the Moderator of this forum
			if chkForumModerator(Request.Form("FORUM_ID"), STRdbntUserName) = "0" then
				Go_Result "只有管理员、版主或作者才能修改此文章", 0
			end if
		case 4 '## Admin so OK
			'## Do Nothing
		case else 
			Go_Result cstr(Member), 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
	end select

	txtMessage = ChkString(Request.Form("Message"),"message")
	Err_Msg = ""

	if txtMessage = " " then 
		Err_Msg = Err_Msg & "<li>不是吧你？内容都不写？</li>"
	end if
	if Err_Msg = "" then
		if strEditedByDate = "1" and mlev < 4 then
			txtMessage = txtMessage & vbCrLf & vbCrLf & "<font size=""" & strFooterFontSize & """ color=""" & strLinkColor & """>Edited by - "
			txtMessage = txtMessage & ChkString(STRdbntUserName, "display") & " 重新编辑於 " & ChkDate(DateToStr(strForumTimeAdjust)) & " " & ChkTime(DateToStr(strForumTimeAdjust)) & "</font>"
		end if

		'## Forum_SQL - Do DB Update
		strSql = "UPDATE " & strTablePrefix & "REPLY "
		strSql = strSql & " SET R_MESSAGE = '" & txtMessage & "'"
		if lcase(strEmail) = "1" then '**
			if Request.Form("rmail") <> "1" then
				TF = "0"
			else 
				TF = "1"
			end if
			strSql = strSql & ", R_MAIL = " & TF
		end if
		strSql = strSql & " WHERE REPLY_ID=" & Request.Form("REPLY_ID")

		my_Conn.Execute (strSql)

		if mLev <> 4 then
			'## Forum_SQL - Update Last Post
			strSql = " UPDATE " & strTablePrefix & "FORUM"
			strSql = strSql & " SET F_LAST_POST = '" & DateToStr(strForumTimeAdjust) & "'"
			strSql = strSql & ",    F_LAST_POST_AUTHOR = " & getMemberID(STRdbntUserName)
			strSql = strSql & " WHERE FORUM_ID = " & Request.Form("FORUM_ID")

			my_Conn.Execute (strSql)
			
			'## Forum_SQL - Update Last Post
			strSql = " UPDATE " & strTablePrefix & "TOPICS"
			strSql = strSql & " SET T_LAST_POST = '" & DateToStr(strForumTimeAdjust) & "'"
			strSql = strSql & ",    T_LAST_POST_AUTHOR = " & getMemberID(STRdbntUserName)
			strSql = strSql & " WHERE TOPIC_ID = " & Request.Form("TOPIC_ID")

			my_Conn.Execute (strSql)
		
		end if

		err_Msg = ""
		if Err.description <> "" then 
			Go_Result "发生一个错误 →  " & Err.description, 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
		else
			Go_Result "更新完成", 1
		end if

		'## Forum_SQL
		strSql = "UPDATE " & strTablePrefix & "TOPICS "
		strSql = strSql & " SET T_LAST_POST = '" & DateToStr(strForumTimeAdjust) & "'"
		strSql = strSql & ",    T_LAST_POST_AUTHOR = " & getMemberID(STRdbntUserName)
		strSql = strSql & " WHERE TOPIC_ID = " & Request.Form("TOPIC_ID")

		my_Conn.Execute (strSql)

		err_Msg = ""
		if Err.description <> "" then 
			Go_Result "发生一个错误 →  " & Err.description, 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
		else
			Go_Result  "更新完成", 1
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
		end if
	else 
%>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">你输入的资料有问题或没有填写完整</font></p>

	<table align=center border=0>
	  <tr>
	    <td><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
		<ul>
		<% =Err_Msg %>
		</ul>
	    </font></td>
	  </tr>
	</table>

	<p align=center><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>><a href="JavaScript:history.go(-1)">请返回重新输入</a></font></p>
<%
	end if
end if

if Request.Form("Method_Type") = "EditTopic" then
	member = cint(ChkUser(STRdbntUserName, Request.Form("Password")))
	select case Member 
		case 0 '## Invalid Pword
			Go_Result "错误的用户名跟密码", 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
		case 1 '## Author of Post so OK
			'## Do Nothing
		case 2 '## Normal User - Not Authorised
			Go_Result "只有管理员、版主或作者才能修改此文章", 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
		case 3 '## Moderator so 
			if chkForumModerator(Request.Form("FORUM_ID"), STRdbntUserName) = "0" then
				Go_Result "只有管理员、版主或作者才能修改此文章", 0
			end if
		case 4 '## Admin so OK
			'## Do Nothing
		case else 
			Go_Result cstr(Member), 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
	end select

	txtMessage = ChkString(Request.Form("Message"),"message")
	txtSubject = ChkString(Request.Form("Subject"),"title")
	Err_Msg = ""

	if txtSubject = " " then 
		Err_Msg = Err_Msg & "<li>你必须输入文章标题</li>"
	end if
	if txtMessage = " " then 
		Err_Msg = Err_Msg & "<li>你必须输入文章内容</li>"
	end if
	if Err_Msg = "" then
		if strEditedByDate = "1" and mlev < 4 then
			txtMessage = txtMessage & vbCrLf & vbCrLf & "<font size=""" & strFooterFontSize & """ color=""" & strLinkColor & """>Edited by - "
			txtMessage = txtMessage & Chkstring(STRdbntUserName, "display") & " on " & ChkDate(DateToStr(strForumTimeAdjust)) & " " & ChkTime(DateToStr(strForumTimeAdjust)) & "</font>"
		end if

		'## Set array to pull out CAT_ID and FORUM_ID from dropdown values in post.asp
		aryForum = split(Request.Form("Forum"), "|")

		'## Forum_SQL
		strSql = "UPDATE " & strTablePrefix & "TOPICS "
		strSql = strSql & " SET T_MESSAGE = '" & txtMessage & "'"
		strSql = strSql & ",     T_SUBJECT = '" & txtSubject & "'"
		if Request.Form("FORUM_ID") <> "" and Request.Form("FORUM_ID") <> aryForum(1) then
			strSql = strSql & ",     CAT_ID = " & aryForum(0)
			strSql = strSql & ",     FORUM_ID = " & aryForum(1)
		end if
		if lcase(strEmail) = "1" then '**
			if Request.Form("rmail") <> "1" then
				TF = "0"
			else 
				TF = "1"
			end if
			strSql = strSql & ",     T_MAIL = " & TF
		end if
		strSql = strSql & " WHERE TOPIC_ID = " & Request.Form("TOPIC_ID")

		my_Conn.Execute(strSql)

		if Request.Form("FORUM_ID") <> aryForum(1) then
			'## Forum_SQL
			strSql = "UPDATE " & strTablePrefix & "REPLY "
			strSql = strSql & " SET CAT_ID = " & aryForum(0)
			strSql = strSql & ",     FORUM_ID = " & aryForum(1)
			strSql = strSql & " WHERE TOPIC_ID = " & Request.Form("TOPIC_ID")

			my_Conn.Execute(strSql)
			
			set rs = Server.CreateObject("ADODB.Recordset")
			
			'## Forum_SQL - count total number of replies in Topics table
			strSql = "SELECT T_REPLIES, T_LAST_POST, T_LAST_POST_AUTHOR "
			strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
			strSql = strSql & " WHERE TOPIC_ID = " & Request.Form("TOPIC_ID")

			set rs = my_Conn.Execute (strSql)
			
			intResetCount = rs("T_REPLIES") + 1
			strT_Last_Post = rs("T_LAST_POST")
			strT_Last_Post_Author = rs("T_LAST_POST_AUTHOR")
			
			rs.Close
			set rs = nothing

			'## Forum_SQL - Get last_post and last_post_author for MoveFrom-Forum
			strSql = "SELECT T_LAST_POST, T_LAST_POST_AUTHOR "
			strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
			strSql = strSql & " WHERE FORUM_ID = " & Request.Form("FORUM_ID") & " "
			strSql = strSql & " ORDER BY T_LAST_POST DESC;"

			set rs = my_Conn.Execute (strSql)
			
			if not rs.eof then
				strLast_Post = rs("T_LAST_POST")
				strLast_Post_Author = rs("T_LAST_POST_AUTHOR")
			else
				strLast_Post = ""
				strLast_Post_Author = ""
			end if
			
			rs.Close
			set rs = nothing

			'## Forum_SQL - Update count of replies to a topic in Forum table



			strSql = "UPDATE " & strTablePrefix & "FORUM SET "
			strSql = strSql & " F_COUNT = F_COUNT - " & intResetCount
			if strLast_Post <> "" then 
				strSql = strSql & ", F_LAST_POST = '" & strLast_Post & "'"
				if strLast_Post_Author <> "" then 
					strSql = strSql & ", F_LAST_POST_AUTHOR = " & strLast_Post_Author
				end if
			end if
			strSql = strSql & " WHERE FORUM_ID = " & Request.Form("FORUM_ID")
			my_Conn.Execute(strSql)

			'## Forum_SQL
			strSql =  "UPDATE " & strTablePrefix & "FORUM SET "
			strSql = strSql & " F_TOPICS = F_TOPICS - 1 "
			strSql = strSql & " WHERE FORUM_ID = " & Request.Form("FORUM_ID")
			my_Conn.Execute(strSql)

			'## Forum_SQL - Get last_post and last_post_author for Forum
			strSql = "SELECT T_LAST_POST, T_LAST_POST_AUTHOR "
			strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
			strSql = strSql & " WHERE FORUM_ID = " & aryForum(1) & " "
			strSql = strSql & " ORDER BY T_LAST_POST DESC;"

			set rs = my_Conn.Execute (strSql)
			
			if not rs.eof then
				strLast_Post = rs("T_LAST_POST")
				strLast_Post_Author = rs("T_LAST_POST_AUTHOR")
			else
				strLast_Post = ""
				strLast_Post_Author = ""
			end if
			
			rs.Close
			set rs = nothing

			'## Forum_SQL - Update count of replies to a topic in Forum table
			strSql = "UPDATE " & strTablePrefix & "FORUM SET "
			strSql = strSql & " F_COUNT = (F_COUNT + " & intResetCount & ")"
			if strLast_Post <> "" then 
				strSql = strSql & ", F_LAST_POST = '" & strLast_Post & "'"
				if strLast_Post_Author <> "" then 
					strSql = strSql & ", F_LAST_POST_AUTHOR = " & strLast_Post_Author
				end if
			end if
			strSql = strSql & " WHERE FORUM_ID = " & aryForum(1)
			my_Conn.Execute(strSql)

			'## Forum_SQL
			strSql =  "UPDATE " & strTablePrefix & "FORUM SET "
			strSql = strSql & " F_TOPICS = F_TOPICS + 1 "
			strSql = strSql & " WHERE FORUM_ID = " & aryForum(1)
			my_Conn.Execute(strSql)

		end if
		err_Msg = ""
		aryForum = ""
		if Err.description <> "" then 
			Go_Result "发生一个错误 →  " & Err.description, 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
		else
			Go_Result  "更新完成", 1
		end if
	else 
%>
	<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">你输入的资料有问题或没有填写完整</font></p>

	<table align=center border=0>
	  <tr>
	    <td><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
		<ul>
		<% =Err_Msg %>
		</ul>
	    </font></td>
	  </tr>
	</table>

	<p align=center><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>><a href="JavaScript:history.go(-1)">请返回重新输入</a></font></p>
<%
	end if
end if

if Request.Form("Method_Type") = "Topic" then
	'## Forum_SQL
	strSql = "SELECT MEMBER_ID, M_LEVEL, M_EMAIL, "&Strdbntsqlname
	if strAuthType = "db" then
		strSql = strSql & ", M_PASSWORD "
	end if
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE "&Strdbntsqlname&" = '" & STRdbntUserName & "'"
	strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_STATUS = " & 1
	if strAuthType = "db" then
		strSql = strSql & " AND   M_PASSWORD = '" & Request.Form("Password") &"'"
		QuoteOk = (ChkQuoteOk(STRdbntUserName) and ChkQuoteOk(Request.Form("Password")))
	else
		QuoteOk = ChkQuoteOk(Session(strCookieURL & "userid"))
	end if

	set rs = my_Conn.Execute (strSql)

	if rs.BOF or rs.EOF or not(QuoteOk) then '##  Invalid Password
		Go_Result "Invalid UserName or Password", 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
		Response.End
	else

			if not(chkForumAccess(Request.Form("FORUM_ID"))) then			
				Go_Result "你不能在本论坛发表文章" & strDBNTUserName, 0		
			end if
		txtMessage = ChkString(Request.Form("Message"),"message")
		txtSubject = ChkString(Request.Form("Subject"),"title")

		if txtMessage = " " then
			Go_Result "你必须填写内容", 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
		end if
		if txtSubject = " " then
			Go_Result "你必须填写标题!", 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
		end if         
		if Request.Form("sig") = "yes" and GetSig(STRdbntUserName) <> "" then
		     txtMessage = txtMessage & vbCrLf & vbCrLf & ChkString(GetSig(STRdbntUserName), "signature" )
		end if
'
		if Request.Form("rmail") <> "1" then
			TF = "0"
		else 
			TF = "1"
		end if

		'## Forum_SQL - Add new post to Topics Table
		strSql = "INSERT INTO " & strTablePrefix & "TOPICS (FORUM_ID"
		strSql = strSql & ", CAT_ID"
		strSql = strSql & ", T_SUBJECT"
		strSql = strSql & ", T_MESSAGE"
		strSql = strSql & ", T_AUTHOR"
		strSql = strSql & ", T_LAST_POST"
		strSql = strSql & ", T_LAST_POST_AUTHOR"
		strSql = strSql & ", T_DATE"
		strSql = strSql & ", T_STATUS"
		if strIPLogging <> "0" then
			strSql = strSql & ", T_IP"
		end if
		strSql = strSql & ", T_MAIL"
		strSql = strSql & ") VALUES ("
		strSql = strSql & Request.Form("FORUM_ID")
		strSql = strSql & ", " & Request.Form("CAT_ID")
		strSql = strSql & ", '" & txtSubject & "'"
		strSql = strSql & ", '" & txtMessage & "'"
		strSql = strSql & ", " & rs("MEMBER_ID")
		strSql = strSql & ", '" & DateToStr(strForumTimeAdjust) & "'"
		strSql = strSql & ", " & rs("MEMBER_ID")
		strSql = strSql & ", '" & DateToStr(strForumTimeAdjust) & "'"
		if Request.Form("lock") = 1 then
			strSql = strSql & ", 0 "
		else
		 	strSql = strSql & ", 1 "
		end if
		if strIPLogging <> "0" then
			strSql = strSql & ", '" & remoteIP() & "'"
		end if
		strSql = strSql & ", " & TF & ")"

		my_Conn.Execute (strSql)

		if Err.description <> "" then 
			err_Msg = "发生一个错误 →  " & Err.description
		else
			err_Msg = "更新完成"
		end if

		'## Forum_SQL - Increase count of topics and replies in Forum table by 1
		strSql = "UPDATE " & strTablePrefix & "FORUM "
		strSql = strSql & " SET F_LAST_POST = '" & DateToStr(strForumTimeAdjust) & "'"
		strSql = strSql & ",    F_TOPICS = F_TOPICS + 1"
		strSql = strSql & ",    F_COUNT = F_COUNT + 1"
		strSql = strSql & ",    F_LAST_POST_AUTHOR = " & rs("MEMBER_ID") & ""
		strSql = strSql & " WHERE FORUM_ID = " & Request.Form("FORUM_ID")

		my_Conn.Execute (strSql)

		Go_Result err_Msg, 1
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
		Response.End
	end if	
end if

if Request.Form("Method_Type") = "Reply" or Request.Form("Method_Type") = "ReplyQuote" or Request.Form("Method_Type") = "TopicQuote" then
	'## Forum_SQL
	strSql = "SELECT MEMBER_ID, M_LEVEL, M_EMAIL, "&Strdbntsqlname
	if strAuthType = "db" then
	strSql = strSql & ", M_PASSWORD "
	end if
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE "&Strdbntsqlname&" = '" & STRdbntUserName & "'"
	strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_STATUS = " & 1
	if strAuthType = "db" then
		strSql = strSql & " AND   M_PASSWORD = '" & Request.Form("Password") &"'"
		QuoteOk = (ChkQuoteOk(STRdbntUserName) and ChkQuoteOk(Request.Form("Password")))
	else
		QuoteOk = ChkQuoteOk(STRdbntUserName)
	end if

	set rs = my_Conn.Execute (strSql)

	if rs.BOF or rs.EOF or not(QuoteOk) then '##  Invalid Password
		err_Msg = "Invalid Password or User Name"
		Go_Result(err_Msg), 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
		Response.End
	else

		if not(chkForumAccess(Request.Form("FORUM_ID"))) then
					Go_Result "你不能在本论坛发表文章"& strDBNTUserName, 0
		end if

		txtMessage = ChkString(Request.Form("Message"),"message")

		if txtMessage = " " then
			Go_Result "你必须填写内容", 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
		end if
		
		if Request.Form("sig") = "yes" and GetSig(STRdbntUserName) <> "" then
		     txtMessage = txtMessage & vbCrLf & vbCrLf & ChkString(GetSig(STRdbntUserName), "signature" )
		end if

		DoReplyEmail Request.Form("TOPIC_ID"), rs("MEMBER_ID"), Request.Form("UserName")

		if Request.Form("rmail") <> "1" then
			RF  = "0"
		else
			RF = "1"
		end if
if CanUserPost(Request.Form("TOPIC_ID"), rs("MEMBER_ID")) = "yes" then
		'## Forum_SQL
		strSql = "INSERT INTO " & strTablePrefix & "REPLY "
		strSql = strSql & "(TOPIC_ID"
		strSql = strSql & ", FORUM_ID"
		strSql = strSql & ", CAT_ID"
		strSql = strSql & ", R_AUTHOR"
		strSql = strSql & ", R_DATE "
		if strIPLogging <> "0" then
			strSql = strSql & ", R_IP"
		end if
		strSql = strSql & ", R_MAIL"
		strSql = strSql & ", R_MESSAGE"
		strSql = strSql & ") VALUES ("
		strSql = strSql & Request.Form("TOPIC_ID")
		strSql = strSql & ", " & Request.Form("FORUM_ID")
		strSql = strSql & ", " & Request.Form("CAT_ID")
		strSql = strSql & ", " & rs("MEMBER_ID")
		strSql = strSql & ", " & "'" & DateToStr(strForumTimeAdjust) & "'"
		if strIPLogging <> "0" then
			strSql = strSql & ", " & "'" & remoteIP() & "'"
		end if
		strSql = strSql & ", " & RF
		strSql = strSql & ", " & "'" & txtMessage & "'"
		strSql = strSql & ")"

		my_Conn.Execute (strSql)

'###################### Attach Files #######################
		strSQL = "SELECT MAX(REPLY_ID) AS MAXID FROM " & strTablePrefix & "REPLY "
		strSQL = strSQL + "WHERE TOPIC_ID = " & Request.Form("TOPIC_ID") 
		strSQL = strSQL + " AND R_AUTHOR = " & rs("MEMBER_ID")
		set rs2 = my_Conn.Execute (strSql)
		intReplyID = rs2("MAXID")
		rs2.close
		strSQL = "UPDATE " & strTablePrefix & "USERFILES "
		strSQL = strSQL & "SET F_REPLY_ID = " & intReplyID 
		strSQL = strSQL & " WHERE MEMBER_ID = " & rs("MEMBER_ID")
		strSQL = strSQL & " AND F_REPLY_ID = -1 AND F_TOPIC_ID =" & Request.Form("TOPIC_ID")
		my_Conn.execute (strSQL)
'###################### Attach Files #######################		

		'## Forum_SQL - Update Last Post and count
		strSql = "UPDATE " & strTablePrefix & "TOPICS "
		strSql = strSql & " SET T_LAST_POST = '" & DateToStr(strForumTimeAdjust) & "'"
		strSql = strSql & ",    T_REPLIES = T_REPLIES + 1 "
		strSql = strSql & ",    T_LAST_POST_AUTHOR = " & rs("MEMBER_ID")
		if Request.Form("lock") = 1 then
			strSql = strSql & ",	T_STATUS = 0 "
		end if
		strSql = strSql & " WHERE TOPIC_ID = " & Request.Form("TOPIC_ID")

		my_Conn.Execute (strSql)

		'## Forum_SQL
		strSql = "UPDATE " & strTablePrefix & "FORUM "
		strSql = strSql & " SET F_LAST_POST = '" & DateToStr(strForumTimeAdjust) & "'"
		strSql = strSql & ",	F_LAST_POST_AUTHOR = " & rs("MEMBER_ID")
		strSql = strSql & ",    F_COUNT = F_COUNT + 1 "
		strSql = strSql & " WHERE FORUM_ID = " & Request.Form("FORUM_ID")

		my_Conn.Execute (strSql)
else
	Err.description = "你已经发表过，必须等60秒后才能重新发表"
end if
		if Err.description <> "" then 
			Go_Result  "发生一个错误 →  " & Err.description, 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
		else
			if Request.Form("M") = "1" then 
				'## Forum_SQL
				strSql  = "SELECT " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.M_EMAIL "
				strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS, " & strTablePrefix & "TOPICS "
				strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "TOPICS.T_AUTHOR "
				strSql = strSql & " AND   " & strTablePrefix & "TOPICS.TOPIC_ID = " & Request.Form("TOPIC_ID")

				set rs2 = my_Conn.Execute (strSql)
				
				DoEmail  rs2("M_EMAIL"), rs2("M_NAME")
				rs2.close
				set rs2 = nothing
			end if
			Go_Result  "更新完成", 1
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
	     end if
	end if
end if

if Request.Form("Method_Type") = "Forum" then
	member = cint(ChkUser(strDBNTUserName, Request.Form("Password")))
	select case Member
		case 0 
			'## Invalid Pword
			Go_Result "错误的用户名跟密码", 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
		case 1 '## Author of Post
			'## Do Nothing
		case 2 '## Normal User - Not Authorised
			Go_Result "只有版主能建立论坛", 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
		case 3 '## Moderator
			if chkForumModerator(Request.Form("FORUM_ID"), STRdbntUserName) = "0" then
				Go_Result "只有版主能建立新论坛", 0
			end if

		case 4 '## Admin
			'## Do Nothing
		case else 
			Go_Result cstr(Member), 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
	end select

	txtMessage = ChkString(Request.Form("Message"),"message")
	txtSubject = ChkString(Request.Form("Subject"),"title")
	Err_Msg = ""

	if txtSubject = " " then 
		Err_Msg = Err_Msg & "<li>你必须输入新论坛的名称</li>"
	end if
	if txtMessage = "" then 
		Err_Msg = Err_Msg & "<li>你必须输入新论坛的简介</li>"
	end if
	if Err_Msg = "" then
		'## Forum_SQL - Do DB Update
		strSql = "INSERT INTO " & strTablePrefix & "FORUM "
		strSql = strSql & "(CAT_ID"
		if strPrivateForums = "1" then
			strSql = strSql & ", F_PRIVATEFORUMS"
			strSql = strSql & ", F_PASSWORD_NEW"
'##########
			strSql = strSql & ", F_HIDDEN"
'##########
'			strSql = strSql & ", F_USERLIST"
		end if
		strSql = strSql & ", F_LAST_POST"
		strSql = strSql & ", F_CREATED"
		strSql = strSql & ", F_SUBJECT"
		strSql = strSql & ", F_DESCRIPTION"
		strSql = strSql & ", F_TYPE" 
		strSql = strSql & ") VALUES ("
		strSql = strSql & Request.Form("CAT_ID")
		if strPrivateForums = "1" then
			strSql = strSql & ", " & Request.Form("AuthType") & ""
			strSql = strSql & ", '" & ChkString(Request.Form("AuthPassword"),"password") & "'"
'##########
			if Request.Form("HideForum") = 1 then
			strSql = strSql & ", 1 "
			else
			strSql = strSql & ", 0 "
			end if
'##########
'			strSql = strSql & ", '" & ChkString(Request.Form("AuthUsers"),"list") & "'"
		end if
		strSql = strSql & ", " & "'" & DateToStr(strForumTimeAdjust) & "'"
		strSql = strSql & ", " & "'" & DateToStr(strForumTimeAdjust) & "'"
		strSql = strSql & ", '" & txtSubject & "'"
		strSql = strSql & ", '" & txtMessage & "'"
		strSql = strSql & ", " & Request.Form("Type")
		strSql = strSql & ")"

		my_Conn.Execute (strSql)

		err_Msg = ""
		if Err.description <> "" then 
			Go_Result "发生一个错误 →  " & Err.description, 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
		Else
'######## Update allowed user list##################################		
			set rsCount = my_Conn.execute("SELECT MAX(FORUM_ID) AS maxForumID FROM " & strTablePrefix & "FORUM ")
			newForumMembers rsCount("maxForumId")
'##################################################################
			Go_Result  "更新完成", 1
		end if
	else 
%>
	<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">你输入的资料有问题或没有填写完整</font></p>

	<table align=center border=0>
	  <tr>
	    <td><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
		<ul>
		<% =Err_Msg %>
		</ul>
	    </font></td>
	  </tr>
	</table>

	<p align=center><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>><a href="JavaScript:history.go(-1)">请返回重新输入</a></font></p>
<%
	end if
end if

if Request.Form("Method_Type") = "URL" then
	member = cint(ChkUser(strDBNTUserName, Request.Form("Password")))
	select case Member
		case 0'## Invalid Pword
			Go_Result "错误的用户名跟密码", 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
		case 1 '## Author of Post
			'## Do Nothing
		case 2 '## Normal User - Not Authorised
			Go_Result "只有版主能建立网路连接", 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
		case 3 '## Moderator
			if chkForumModerator(Request.Form("FORUM_ID"), STRdbntUserName) = "0" then
				Go_Result "只有版主能建立网路连接", 0
			end if
		case 4 '## Admin
			'## Do Nothing
		case else 
			Go_Result cstr(Member), 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
	end select

	txtMessage = ChkString(Request.Form("Message"),"message")
	txtAddress = ChkString(Request.Form("Address"),"url")
	txtSubject = ChkString(Request.Form("Subject"),"title")
	Err_Msg = ""

	if txtSubject = " " then 
		Err_Msg = Err_Msg & "<li>你必须输入新连接的名称</li>"
	end if
	if txtAddress = " " or lcase(txtAddress) = "http://" or lcase(txtAddress) = "https://" or lcase(txtAddress) = "file:///" then 
		Err_Msg = Err_Msg & "<li>你必须输入新连接的网址</li>"
	end if
	if (left(lcase(txtAddress), 7) <> "http://" and left(lcase(txtAddress), 8) <> "https://") and left(lcase(txtAddress), 8) <> "file:///" and txtAddress <> "" then
		Err_Msg = Err_Msg & "<li>你必须在网址前加上 <b>http://</b>, <b>https://</b> or <b>file:///</b></li>"
	end if
	if txtMessage = " " then 
		Err_Msg = Err_Msg & "<li>你必须输入新连接的简介</li>"
	end if
	if Err_Msg = "" then
		'## Forum_SQL - Do DB Update
		strSql = "INSERT INTO " & strTablePrefix & "FORUM "
		strSql = strSql & "(CAT_ID"
		if strPrivateForums = "1" then
			strSql = strSql & ", F_PRIVATEFORUMS"
'##########
			strSql = strSql & ", F_HIDDEN"
'##########
'			strSql = strSql & ", F_USERLIST"
		end if
		strSql = strSql & ", F_LAST_POST"
		strSql = strSql & ", F_LAST_POST_AUTHOR"
		strSql = strSql & ", F_SUBJECT"
		strSql = strSql & ", F_URL"
		strSql = strSql & ", F_URLIMAGE"
		strSql = strSql & ", F_DESCRIPTION"
		strSql = strSql & ", F_TYPE"
		strSql = strSql & ")  VALUES ("
		strSql = strSql & Request.Form("CAT_ID")
		if strPrivateForums = "1" then
			strSql = strSql & ", " & Request.Form("AuthType") & ""
'##########
			if Request.Form("HideForum") = 1 then
			strSql = strSql & ", 1 "
			else
			strSql = strSql & ", 0 "
			end if
'##########

'			strSql = strSql & ", '" & ChkString(Request.Form("AuthUsers"),"list") & "'"
		end if
		strSql = strSql & ", " & "'" & DateToStr(strForumTimeAdjust) & "'"
		strSql = strSql & ", " & getMemberID(Request.Form("UserName")) & " "
		strSql = strSql & ", " & "'" & txtSubject & "'"
		strSql = strSql & ", " & "'" & txtAddress & "'"
		strSql = strSql & ", " & "'" & Request.Form("UrlImage") & "'"
		strSql = strSql & ", " & "'" & txtMessage & "'"
		strSql = strSql & ", " & Request.Form("Type")
		strSql = strSql & ") "

		my_Conn.Execute (strSql)

		err_Msg = ""
		if Err.description <> "" then 
			Go_Result "发生一个错误 →  " & Err.description, 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
		else
'########### Update allowed user list ##############################
			set rsCount = my_Conn.execute("SELECT MAX(FORUM_ID) AS maxForumID FROM " & strTablePrefix & "FORUM ")
			newForumMembers rsCount("maxForumId")                   
'##################################################################
			Go_Result  "更新完成", 1
		end if
	else 
%>
	<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">你输入的资料有问题或没有填写完整</font></p>

	<table align=center border=0>
	  <tr>
	    <td><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
		<ul>
		<% =Err_Msg %>
		</ul>
	    </font></td>
	  </tr>
	</table>

	<p align=center><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>><a href="JavaScript:history.go(-1)">请返回重新输入</a></font></p>
<%
	end if
end if

if Request.Form("Method_Type") = "EditForum" then
	member = cint(ChkUser(STRdbntUserName, Request.Form("Password")))
	select case Member 
		case 0 
			'## Invalid Pword
			Go_Result "错误的用户名跟密码", 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
		case 1 '## Author of Post
			 '## Do Nothing
		case 2 '## Normal User - Not Authorised
			Go_Result "只有管理员和版主可以修改论坛属性", 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
		case 3 '## Moderator
			if chkForumModerator(Request.Form("FORUM_ID"), STRdbntUserName) = "0" then
				Go_Result "Only Admins and Moderators change this Forum", 0
			end if	
		case 4 '## Admin
			'## Do Nothing
		case else 
			Go_Result cstr(Member), 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
	end select

	txtMessage = ChkString(Request.Form("Message"),"message")
	txtSubject = ChkString(Request.Form("Subject"),"title")
	Err_Msg = ""

	if txtSubject = " " then 
		Err_Msg = Err_Msg & "<li>你必须输入论坛的名称</li>"
	end if
	if txtMessage = " " then 
		Err_Msg = Err_Msg & "<li>你必须输入论坛的简介</li>"
	end if
	if Err_Msg = "" then
		'## Forum_SQL - Do DB Update
		strSql = "UPDATE " & strTablePrefix & "FORUM "
		strSql = strSql & " SET CAT_ID = " & Request.Form("Category")
		if strPrivateForums = "1" then
			strSql = strSql & ",    F_PRIVATEFORUMS = " & Request.Form("AuthType") & ""
			strSql = strSql & ",    F_PASSWORD_NEW = '" & ChkString(Request.Form("AuthPassword"),"password") & "'"
'###########
			if Request.Form("HideForum") = 1 then
			strSql = strSql & ", F_HIDDEN = 1 "
			else
			strSql = strSql & ", F_HIDDEN = 0 "
			end if
'###########
'			strSql = strSql & ",    F_USERLIST = '" & ChkString(Request.Form("AuthUsers"),"list") & "'"
		end if
		strSql = strSql & ",    F_SUBJECT = '" & txtSubject & "'"
		strSql = strSql & ",    F_DESCRIPTION = '" & txtMessage & "'"
		strSql = strSql & " WHERE FORUM_ID = " & Request.Form("FORUM_ID")

		my_Conn.Execute (strSql)

		 err_Msg= ""
		if Err.description <> "" then 
			Go_Result "发生一个错误 →  " & Err.description, 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
		else
'########## Update Allowed user List ###############################
			set rsCount = my_Conn.execute("SELECT MAX(FORUM_ID) AS maxForumID FROM " & strTablePrefix & "FORUM ")
			updateForumMembers Request.Form("FORUM_ID")           
'###################################################################
			Go_Result  "更新完成", 1
		end if
	else 
%>
	<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">你输入的资料有问题或没有填写完整</font></p>

	<table align=center border=0>
	  <tr>
	    <td><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
		<ul>
		<% =Err_Msg %>
		</ul>
	    </font></td>
	  </tr>
	</table>

	<p align=center><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>><a href="JavaScript:history.go(-1)">请返回重新输入</a></font></p>
<%
	end if
end if

if Request.Form("Method_Type") = "EditURL" then
	member = cint(ChkUser(strDBNTUserName, Request.Form("Password")))
	select case Member 
		case 0 
			'## Invalid Pword
			Go_Result "错误的用户名跟密码", 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
		case 1 '## Author of Post
			 '## Do Nothing
		case 2 '## Normal User - Not Authorised
			Go_Result "只有管理员和版主可以修改论坛属性", 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
		case 3 '## Moderator
			if chkForumModerator(Request.Form("FORUM_ID"), STRdbntUserName) = "0" then
				Go_Result "Only Admins and Moderators change this Forum", 0
			end if	
		case 4 '## Admin
			'## Do Nothing
		case else 
			Go_Result cstr(Member), 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
	end select

	txtMessage = ChkString(Request.Form("Message"),"message")
	txtAddress = ChkString(Request.Form("Address"),"url")
	txtSubject = ChkString(Request.Form("Subject"),"title")
	Err_Msg = ""

	if txtSubject = " " then 
		Err_Msg = Err_Msg & "<li>你必须输入新连接的名称</li>"
	end if
	if txtAddress = " " or lcase(txtAddress) = "http://" or lcase(txtAddress) = "https://" or lcase(txtAddress) = "file:///" then 
		Err_Msg = Err_Msg & "<li>你必须输入新连接的网址</li>"
	end if
	if (left(lcase(txtAddress), 7) <> "http://" and left(lcase(txtAddress), 8) <> "https://" and left(lcase(txtAddress), 8) <> "file:///") and (txtAddress <> "") then
		Err_Msg = Err_Msg & "<li>你必须在网址前加上 <b>http://</b>, <b>https://</b> or <b>file:///</b></li>"
	end if
	if txtMessage = "" then 
		Err_Msg = Err_Msg & "<li>你必须输入新连接的简介</li>"
	end if
	if Err_Msg = "" then

		'## Forum_SQL - Do DB Update
		strSql = "UPDATE " & strTablePrefix & "FORUM "
		strSql = strSql & " SET CAT_ID = " & Request.Form("Category")
		if strPrivateForums = "1" then
			strSql = strSql & ",    F_PRIVATEFORUMS = " & Request.Form("AuthType") & ""
'#############
			if Request.Form("HideForum") = 1 then
			strSql = strSql & ", F_HIDDEN=1 "
			else
			strSql = strSql & ", F_HIDDEN=0 "
			end if
'#############
'			strSql = strSql & ",    F_USERLIST = '" & ChkString(Request.Form("AuthUsers"),"list") & "'"
		end if
		strSql = strSql & ",    F_SUBJECT = '" & txtSubject & "'"
		strSql = strSql & ",    F_URL = '" & txtAddress & "'"
		strSql = strSql & ",    F_URLIMAGE = '" & Request.Form("UrlImage") & "'"
		strSql = strSql & ",    F_DESCRIPTION = '" & txtMessage & "'"

		strSql = strSql & " WHERE FORUM_ID = " & Request.Form("FORUM_ID")

		my_Conn.Execute (strSql)

		 err_Msg= ""
		if Err.description <> "" then 
			Go_Result "发生一个错误 →  " & Err.description, 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
		else
'########## Update Allowed user List ###############################
			set rsCount = my_Conn.execute("SELECT MAX(FORUM_ID) AS maxForumID FROM " & strTablePrefix & "FORUM ")
			updateForumMembers Request.Form("FORUM_ID")
'###################################################################
			Go_Result  "更新完成", 1
		end if
	else 
%>
	<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">你输入的资料有问题或没有填写完整</font></p>

	<table align=center border=0>
	  <tr>
	    <td><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
		<ul>
		<% =Err_Msg %>
		</ul>
	    </font></td>
	  </tr>
	</table>

	<p align=center><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>><a href="JavaScript:history.go(-1)">请返回重新输入</a></font></p>
<%
	end if
end if

if Request.Form("Method_Type") = "Category" then
	member = cint(ChkUser(STRdbntUserName, Request.Form("Password")))
	select case Member 
		case 0 
			'## Invalid Pword
			Go_Result "错误的用户名跟密码", 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
		case 1 '## Author of Post
			'## Do Nothing
		case 2 '## Normal User - Not Authorised
			Go_Result "只有管理员才能建立新分类", 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
		case 3 '## Moderator
			if chkForumModerator(Request.Form("FORUM_ID"), STRdbntUserName) = "0" then
				Go_Result "只有管理员才能建立新分类", 0
			end if	
		case 4 '## Admin
			'## Do Nothing
		case else 
			Go_Result cstr(Member), 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
	end select
	Err_Msg = ""
	if Request.Form("Subject") = "" then 
		Err_Msg = Err_Msg & "<li>你必须输入新分类的名称</li>"
	end if
	if Err_Msg = "" then

		'## Forum_SQL - Do DB Update
		strSql = "INSERT INTO " & strTablePrefix & "CATEGORY (CAT_NAME) "
		strSql = strSql & " VALUES ('" & ChkString(Request.Form("Subject"),"title") & "')"

		my_Conn.Execute (strSql)

		 err_Msg= ""
		if Err.description <> "" then 
			Go_Result "发生一个错误 →  " & Err.description, 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
		else
			Go_Result  "更新完成", 1
		end if
	else 
%>
	<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">你输入的资料有问题或没有填写完整</font></p>

	<table align=center border=0>
	  <tr>
	    <td><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
		<ul>
		<% =Err_Msg %>
		</ul>
	    </font></td>
	  </tr>
	</table>

	<p align=center><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>><a href="JavaScript:history.go(-1)">请返回重新输入</a></font></p>
<%
	end if
end if

if Request.Form("Method_Type") = "EditCategory" then
	member = cint(ChkUser(STRdbntUserName, Request.Form("Password")))
	select case Member 
		case 0 
			'## Invalid Pword
			Go_Result "错误的用户名跟密码", 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
		case 1 '## Author of Post
			'## Do Nothing
		case 2 '## Normal User - Not Authorised
			Go_Result "只有管理员才能修改分类", 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
		case 3 '## Moderator
			'## Do Nothing
			if chkForumModerator(Request.Form("FORUM_ID"), STRdbntUserName) = "0" then
				Go_Result "只有管理员才能修改分类", 0
			end if
		case 4 '## Admin
			'## Do Nothing
		case else 
			Go_Result cstr(Member), 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
	end select
	Err_Msg = ""
	if Request.Form("Subject") = "" then 
		Err_Msg = Err_Msg & "<li>必须输入分类的名称</li>"
	end if
	if Err_Msg = "" then
		'## Forum_SQL - Do DB Update
		strSql = "UPDATE " & strTablePrefix & "CATEGORY "
		strSql = strSql & " SET CAT_NAME = '" & ChkString(Request.Form("Subject"),"title") & "'"
		strSql = strSql & " WHERE CAT_ID = " & Request.Form("CAT_ID")

		my_Conn.Execute (strSql)

		 err_Msg= ""
		if Err.description <> "" then 
			Go_Result "发生一个错误 →  " & Err.description, 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
		End If
			i = 1
			do until i > cint(Request.Form("NumberForums"))
				SelectName = "SortForum" & i 
				SelectID   = "SortForumID" & i
		
				'## Forum_SQL - Do DB Update
				strSql = "UPDATE " & strTablePrefix & "FORUM "
				strSql = strSql & " SET FORUM_ORDER = " & Request.Form(SelectName)
				strSql = strSql & " WHERE FORUM_ID = " & Request.Form(SelectId)
				my_Conn.Execute (strSql)
	
				err_Msg= ""
				if Err.description <> "" then 
					Go_Result "发生一个错误 →  " & Err.description, 0
				End if
				i = i + 1
			loop
		 	Go_Result "更新完成", 1
	else 
%>
	<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">你输入的资料有问题或遗漏</font></p>

	<table align=center border=0>
	  <tr>
	    <td><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
		<ul>
		<% =Err_Msg %>
		</ul>
	    </font></td>
	  </tr>
	</table>

	<p align=center><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>><a href="JavaScript:history.go(-1)">请返回重新输入</a></font></p>
<%
	end if
end if

if Request.Form("Method_Type") = "SortCategory" then
	member = cint(ChkUser(STRdbntUserName, Request.Form("Password")))
	Select Case Member 
		case 0 	'## Invalid Pword
			Go_Result "错误的密码或使用者名称", 0
		case 1 '## Author of Post
			Go_Result "只有管理员可以变更分类排序", 0
		case 2 '## Normal User - Not Authorised
			Go_Result "只有管理员可以变更分类排序", 0
		case 3 '## Moderator
			'## Do Nothing
			if chkForumModerator(Request.Form("FORUM_ID"), STRdbntUserName) = "0" then
				Go_Result "只有管理员可以修改分类排序", 0
			end if
		case 4 '## Admin
			'## Do Nothing
		case Else 
			Go_Result cstr(Member), 0
	End Select

	i = 1
	do until i > cint(Request.Form("NumberCategories"))
		SelectName = "SortCategory" & i 
		SelectID   = "SortCatID" & i
		
		'## Forum_SQL - Do DB Update
		strSql = "UPDATE " & strTablePrefix & "CATEGORY "
		strSql = strSql & " SET CAT_ORDER = " & Request.Form(SelectName)
		strSql = strSql & " WHERE CAT_ID = " & Request.Form(SelectId)
		my_Conn.Execute (strSql)
                
               		
		err_Msg= ""
		if Err.description <> "" then 
			Go_Result "发生一个错误 →  " & Err.description, 0
		End if
		i = i + 1
	loop
 	Go_Result "更新完成", 1
End if

%>
<% set rs = nothing %>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
sub DoEmail(email, user_name)
	'## Emails Topic Author if Requested.  
	'## This needs to be Edited to use your own email component
	'## if you don't have one, try the w3Jmail component from www.dimac.net it's free!
    '#-#-#-#-#-#-#-# Added by Me
    strsql = "Select m_message, m_subject from "&strTablePrefix&"email_config where m_type='newreply'"
    set EmailRS = my_conn.execute(strsql)
    	strEmailSubject = EmailRS("m_subject")
        strEmailMessage = EmailRS("m_message")
    set EmailRS = nothing
    '#-#-#-#-#-#-#-# Done adding
    
    '#-#-#-#-#-#-#-# Changed Code
	if lcase(strEmail) = "1" then
		strRecipientsName = user_name
		strRecipients = email
        
        strSubject = strEmailSubject
        strSubject = Replace(strSubject, "[forum_title]", strForumTitle)
        
        strMessage = strEmailMessage
        strMessage = Replace(strMessage, "[user_name]", user_name)
        strMessage = Replace(strMessage, "[topic_title]", Request.form("Topic_Title"))
        strMessage = Replace(strMessage, "[forum_title]", strForumTitle)
        strMessage = Replace(strMessage, "[link]", Request.form("Refer"))
        strMessage = Replace(strMessage, Chr(10), vbcrlf)
        strMessage = Replace(strMessage, "<br>", vbcrlf)
        strMessage = Replace(strMessage, "<p>", vbcrlf & vbcrlf)
    '#-#-#-#-#-#-#-# End of Changed Code
    '         newreply
    '<<forum_title>> = strForumTitle
    '<<user_name>> = user_name
    '<<topic_title>> = request.form("Topic_Title")
    '<<link>> = Request.form("Refer")
    '#-#-#-#-#-#-#-# Old Code
'	if lcase(strEmail) = "1" then
'		strRecipientsName = user_name
'		strRecipients = email
'		strSubject = strForumTitle & " - Reply to your posting"
'		strMessage = "Hello " & user_name & vbCrLf & vbCrLf
'		strMessage = strMessage & "You have received a reply to your posting on " & strForumTitle & "." & vbCrLf
'		strMessage = strMessage & "Regarding the subject - " & Request.Form("Topic_Title") & "." & vbCrLf & vbCrLf
'		strMessage = strMessage & "You can view the reply at " & Request.Form("Refer") & vbCrLf
%>
<!--#INCLUDE FILE="inc_mail.asp" -->
<%
	end if
end sub

sub DoReplyEmail(TopicNum, PostedBy, PostedByName)
	'## Emails all users who wish to receive a mail if topic
	'## has a reply but only send one per member.
	'## Forum_SQL
	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_EMAIL "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS, " & strTablePrefix & "REPLY "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "REPLY.R_AUTHOR "
	strSql = strSql & " AND   TOPIC_ID = " & TopicNum 
	strSql = strSql & " AND   R_MAIL = 1 "
	strSql = strSql & " ORDER BY " & strMemberTablePrefix & "MEMBERS.MEMBER_ID"

	set rsReply = my_Conn.Execute (strSql)

	'## Forum_SQL
	strSql = " SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.M_EMAIL, " & strTablePrefix & "TOPICS.T_MAIL "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS,  "
	strSql = strSql & strTablePrefix & "TOPICS "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "TOPICS.T_AUTHOR "
	strSql = strSql & " AND " & strTablePrefix & "TOPICS.TOPIC_ID = " & TopicNum

	set rsTopicAuthor = my_Conn.Execute (strSql)
	
	MailSendToAuthor = false
	if (rsTopicAuthor("T_MAIL") = 1) and (PostedBy <> rsTopicAuthor("MEMBER_ID")) then
		strRecipientsName = rsTopicAuthor("M_NAME")
		strRecipients = rsTopicAuthor("M_EMAIL")
        '#-# Added Code
        'replynot
        strsql = "Select m_message, m_subject from "&strTablePrefix&"email_config where m_type='replynot'"
        set EmailRS = my_conn.execute(strsql)
        	strEmailMessage = EmailRS("m_message")
            strEmailSubject = EmailRS("m_subject")
        set EmailRS = nothing
       	'#-# End Added Code
        '#-# Changed Code
        strSubject = strEmailSubject
        strSubject = Replace(strSubject, "[forum_title]", strForumTitle)
        
        strEmailLink = Left(Request.Form("refer"), InstrRev(Request.Form("refer"), "/")) & "link.asp?TOPIC_ID=" & TopicNum & vbCrLf
        
        strMessage = strEmailMessage
        strMessage = Replace(strMessage, "[user_name]", user_name)
        strMessage = Replace(strMessage, "[topic_title]", Request.form("Topic_Title"))
        strMessage = Replace(strMessage, "[forum_title]", strForumTitle)
        strMessage = Replace(strMessage, "[link]", strEmailLink)
        strMessage = Replace(strMessage, "[posted_by]", PostedByName)
        strMessage = Replace(strMessage, Chr(10), vbcrlf)
        strMessage = Replace(strMessage, "<br>", vbcrlf)
        strMessage = Replace(strMessage, "<p>", vbcrlf & vbcrlf)
        '#-# End Changed Code
        
        '#-# Old Code
'		strSubject = strForumTitle & " - Reply to a posting"
'		strMessage = "Hello " & rsTopicAuthor("M_NAME") & vbCrLf & vbCrLf
'		strMessage = strMessage & PostedByName & " has replied to a topic on " & strForumTitle & " that you requested notification to. "
'		strMessage = strMessage & "Regarding the subject - " & Request.Form("Topic_Title") & "." & vbCrLf & vbCrLf
'		strMessage = strMessage & "You can view the reply at " & Left(Request.Form("refer"), InstrRev(Request.Form("refer"), "/")) & "link.asp?TOPIC_ID=" & TopicNum & vbCrLf
%>
<!--#INCLUDE FILE="inc_mail.asp" -->
<%
		MailSendToAuthor = true
	end if
	
	prevMember = ""
	
	do while (not rsReply.EOF) and (not rsReply.BOF)
		if (prevMember <> rsReply("MEMBER_ID")) and (PostedBy <> rsReply("MEMBER_ID")) then
			if (rsTopicAuthor("MEMBER_ID") = rsReply("MEMBER_ID")) and (MailSendToAuthor) then
				'## Do Nothing
				'## The reply was done by the author, and he/she allready has got a mail
			else
				if (rsTopicAuthor("MEMBER_ID") = rsReply("MEMBER_ID")) then
					MailSendToAuthor = true
				end if
				strRecipientsName = rsReply("M_Name")
				strRecipients = rsReply("M_EMAIL")
                '#-# Added Code
        strsql = "Select m_message, m_subject from "&strTablePrefix&"email_config where m_type='replynot'"
        set EmailRS = my_conn.execute(strsql)
        	strEmailMessage = EmailRS("m_message")
            strEmailSubject = EmailRS("m_subject")
        set EmailRS = nothing
                '#-# End Added Code
                '#-# Changed Code
        strSubject = strEmailSubject
        strSubject = Replace(strSubject, "[forum_title]", strForumTitle)
        
        strEmailLink = Left(Request.Form("refer"), InstrRev(Request.Form("refer"), "/")) & "link.asp?TOPIC_ID=" & TopicNum & vbCrLf
        
        strMessage = strEmailMessage
        strMessage = Replace(strMessage, "[user_name]", user_name)
        strMessage = Replace(strMessage, "[topic_title]", Request.form("Topic_Title"))
        strMessage = Replace(strMessage, "[forum_title]", strForumTitle)
        strMessage = Replace(strMessage, "[link]", strEmailLink)
        strMessage = Replace(strMessage, "[posted_by]", PostedByName)
        strMessage = Replace(strMessage, Chr(10), vbcrlf)
        strMessage = Replace(strMessage, "<br>", vbcrlf)
        strMessage = Replace(strMessage, "<p>", vbcrlf & vbcrlf)
                '#-# End Changed Code
                '#-#-# Old Code
                
'				strSubject = strForumTitle & " - Reply to a posting"
'				strMessage = "Hello " & rsReply("M_NAME") & vbCrLf & vbCrLf
'				strMessage = strMessage & PostedByName & " has replied to a topic on " & strForumTitle & " that you requested notification to. "
'				strMessage = strMessage & "Regarding the subject - " & Request.Form("Topic_Title") & "." & vbCrLf & vbCrLf
'				strMessage = strMessage & "You can view the reply at " & Left(Request.Form("refer"), InstrRev(Request.Form("refer"), "/")) & "link.asp?TOPIC_ID=" & TopicNum & vbCrLf
'		'		strMessage = strMessage & "You can view the reply at " & Request.Form("refer") & vbCrLf
%>
<!--#INCLUDE FILE="inc_mail.asp" -->
<%
			end if
		end if
		prevMember = rsReply("MEMBER_ID")
		rsReply.MoveNext
	loop

	rsReply.Close
	set rsReply = nothing

	rsTopicAuthor.Close
	set rsTopicAuthor = nothing
end sub

sub Go_Result(str_err_Msg, boolOk)
%>
<table border="0" width="100%">
  <tr>
	<td width="33%" align="left"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
	<img src="<%=strImageURL %>icon_folder_open.gif" border="0">&nbsp;<a href="default.asp">返回论坛首页</a><br>
<% 
	if Request.form("Method_Type") = "Topic" or _
		Request.form("Method_Type") = "Reply" or _
		Request.form("Method_Type") = "EditTopic" then 
%>
	<img src="<%=strImageURL %>icon_bar.gif" border="0"><img src="<%=strImageURL %>icon_folder_open.gif" border="0">&nbsp;<a href="FORUM.asp?FORUM_ID=<% = Request.Form("FORUM_ID") %>&CAT_ID=<% = Request.Form("CAT_ID") %>&Forum_Title=<% = ChkString(Request.Form("FORUM_Title"),"urlpath")%>"><% = Request.Form("FORUM_Title") %></a><br>
<% 
	end if 
	if Request.form("Method_Type") = "Reply" or _
		Request.form("Method_Type") = "EditTopic" then 
%>
	<img src="<%=strImageURL %>icon_blank.gif" border="0"><img src="<%=strImageURL %>icon_bar.gif" border="0"><img src="<%=strImageURL %>icon_folder_open_topic.gif" border="0">&nbsp;<a href="<% =Request.Form("refer") %>"><% =Request.Form("Topic_Title")%></a>
<% 
	end if 
%>
    </font></td>
  </tr>
</table>

<%
	if boolOk = 1 then 
%>
<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">
<%
		select case Request.Form("Method_Type")
			case "Edit"
				Response.Write("你的回复修改成功！")
			case "EditCategory"
				Response.Write("分类名称修改成功！")
			case "EditForum"
				Response.Write("论坛属性更新成功！")
			case "EditTopic"
				Response.Write("主题修改成功！")
			case "EditURL"
				Response.Write("连接属性更新成功！")
			case "Reply"
				Response.Write("新回复发表完成！")
				DoPCount
				DoUCount Request.Form("UserName")
				DoULastPost Request.Form("UserName")
			case "ReplyQuote"
				Response.Write("新回复发表完成！")
				DoPCount
				DoUCount Request.Form("UserName")
				DoULastPost Request.Form("UserName")
			case "TopicQuote"
				Response.Write("新回复发表完成！")
				DoPCount
				DoUCount Request.Form("UserName")
				DoULastPost Request.Form("UserName")
			case "Topic"
				DoTCount
				DoPCount
				DoUCount Request.Form("UserName")
				DoULastPost Request.Form("UserName")
				Response.Write("新主题发表完成！")
			case "Forum"
				Response.Write("新论坛建立完成！")
			case "URL"
				Response.Write("新连接建立完成！")
			case "Category"
				Response.Write("新分类建立完成！")
			case else
				Response.Write("完成！")
				DoPCount
				DoUCount Request.Form("UserName")
				DoULastPost Request.Form("UserName")
		end select
%>
</font></p>
<meta http-equiv="Refresh" content="0; URL=<% =Request.Form("refer")%>">

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">
<%
		select case Request.Form("Method_Type")
			case "Category"
				Response.Write("记得在本分类建立至少一个讨论区。")
			case "EditCategory"
				Response.Write("感谢你的参与！")
			case "Forum"
				Response.Write("新论坛已经准备完成，可供会员发表文章！")
			case "EditForum"
				Response.Write("感谢你的参与！")
			case "URL"
				Response.Write("新连接已经建立！")
			case "EditURL"
				Response.Write("祝你有愉快的一天！")
			case "Topic"
				Response.Write("感谢你的参与！")
			case "TopicQuote"
				Response.Write("感谢你的参与！")
			case "EditTopic"
				Response.Write("感谢你的参与！")
			case "Reply"
				Response.Write("感谢你的参与！")
			case "ReplyQuote"
				Response.Write("感谢你的参与！")
			case "Edit"
				Response.Write("感谢你的参与！")
			case else
				Response.Write("祝你有愉快的一天！")
		end select
%>
</font></p>

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="<% =Request.Form("refer")%>">返回论坛</font></a></p>

<!--#INCLUDE FILE="inc_footer.asp" -->
<%
		Response.End
	else 
%>
<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">发生问题！</font></p>

<p align="center"><font color="red" size="<% =strHeaderFontSize %>"><% =str_err_Msg %></font></p>

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript:history.go(-1)">返回修正问题。</a></font></p>

<!--#INCLUDE FILE="inc_footer.asp" -->
<%
		Response.End
	end if 
end sub

sub newForumMembers(fForumID)
		on error resume next
		if Request.Form("AuthUsers") = "" then
			exit Sub
		end if
	Users = split(Request.Form("AuthUsers"),",")

	for count = Lbound(Users) to Ubound(Users)
		if Trim(Users(count)) <> "" then
			strSql = "INSERT INTO " & strTablePrefix & "ALLOWED_MEMBERS ("
			strSql = strSql & " MEMBER_ID, FORUM_ID) VALUES ( "& Users(count) & ", " & fForumID & ")"
	
			my_conn.execute (strSql)
			if err.number <> 0 then
				Go_REsult err.description, 0
			end if
		end if
	next

end sub

sub updateForumMembers(fForumID)
		my_Conn.execute ("DELETE FROM " & strTablePrefix & "ALLOWED_MEMBERS WHERE FORUM_ID = " & fForumId)
		newForumMembers(fForumID)
end sub
function CanUserPost(fTopicID, fMemberID)
		lastAllowedPost = DateToStr(DateAdd("n",-1,strForumTimeAdjust))
		strSql = "SELECT MAX(R_DATE) AS LASTPOST FROM " & strTablePrefix & "REPLY WHERE "
		strSql = strSql & "TOPIC_ID = " & fTopicID & " AND R_AUTHOR = " & fMemberID
		rsTmp = my_conn.execute(strSQL)
		if rsTmp("LASTPOST") > lastAllowedPost then
			CanUserPost = "no"
		else
			CanUserPost = "yes"
		end if
end function
%>
