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
<!--#include file="config.asp"-->
<!--#include file="inc_functions.asp"-->
<!--#INCLUDE FILE="inc_top.asp" -->
<%
if strDBNTUserName = "" then Response.redirect ("default.asp")
Select Case Request.QueryString("mode")

Case "deleteReply"
	delRequest = split(Request.Form("delRequest"), ",")
	for i = 0 to ubound(delRequest)
		delSQL = "UPDATE "& strTablePrefix & "REPLY SET R_MAIL=0 WHERE REPLY_ID = " & cint(delRequest(i))
    	my_conn.Execute delSQL
	next
	response.write "<div align=""center""><font face=""" & strDefaultFontFace & """ size=3><b>所选要求已清除！</b></font>"
	response.write "<br><a href=""default.asp"">返回论坛首页</a></div>"
	response.write "<meta http-equiv=""Refresh"" content=""0; URL=cancel_replies.asp"">"
	Response.End

Case "deleteAllReply"
	delSQL = "UPDATE "& strTablePrefix & "REPLY SET R_MAIL = 0 WHERE REPLY_ID IN ("
	delSQL = delSQL & "SELECT REPLY_ID FROM "& strTablePrefix & "REPLY "
	delSQL = delSQL & "WHERE "& strTablePrefix & "REPLY.R_AUTHOR=" & getmemberID(strDBNTUserName) & " AND R_MAIL=1)"
    my_conn.Execute delSQL
	response.write "<center><font face=""" & strDefaultFontFace & """ size=3><b>所有要求皆已清除！</b></font></center>"
	response.write "<div align=""center""><a href=""default.asp"">返回论坛首页</a></div>"
	response.write "<meta http-equiv=""Refresh"" content=""0; URL=cancel_replies.asp"">"
	Response.End

Case "deleteTopic"
	delRequest = split(Request.Form("delRequest"), ",")
	for i = 0 to ubound(delRequest)
		delSQL = "UPDATE "& strTablePrefix & "TOPICS SET T_MAIL=0 WHERE TOPIC_ID = " & cint(delRequest(i))
	    my_conn.Execute delSQL
	next
	response.write "<center><font face=""" & strDefaultFontFace & """ size=3><b>所选要求已清除！</b></font></center>"
	response.write "<div align=""center""><a href=""default.asp"">返回论坛首页</a></div>"
	response.write "<meta http-equiv=""Refresh"" content=""0; URL=cancel_replies.asp"">"
	Response.End

Case "deleteAllTopic"
	delSQL = "UPDATE "& strTablePrefix & "TOPICS SET T_MAIL = 0 WHERE TOPIC_ID IN ("
	delSQL = delSQL & "SELECT TOPIC_ID FROM "& strTablePrefix & "TOPICS "
	delSQL = delSQL & "WHERE "& strTablePrefix & "TOPICS.T_AUTHOR=" & getmemberID(strDBNTUserName) & " AND T_MAIL=1)"
    my_conn.Execute delSQL
	response.write "<center><font face=""" & strDefaultFontFace & """ size=3><b>所选贴子皆已清除！</b></font></center>"
	response.write "<div align=""center""><a href=""default.asp"">返回论坛首页</a></div>"
	response.write "<meta http-equiv=""Refresh"" content=""0; URL=cancel_replies.asp"">"
	Response.End

Case default
	bkReplyPageSize = 5
	bkTopicPageSize = 5
'################################## REPLIES ############
	set rsReply = server.CreateObject("ADODB.RecordSet")
	strSql = "SELECT * FROM "& strTablePrefix & "REPLY "
	strSql = strSql & "WHERE "& strTablePrefix & "REPLY.R_AUTHOR=" & getmemberID(strDBNTUserName)
	strSQL = strSQL & " AND R_MAIL=1"
	rsReply.Open strSQL, my_Conn, 3
	rPage = CLng(request("rPage"))
	rsReply.PageSize = bkReplyPageSize
	rPageCount = cInt(rsReply.PageCount)
	if rPageCount = 0 then rPageCount = 1
	if rPage < 1 then
		rPage = 1
	end if
%>


<form Action="cancel_replies.asp?mode=deleteReply" method=post id=form1 name=form1>
<TABLE align="center" border="0" cellPadding=3 cellSpacing=1 width=85% bgcolor="<%= strTableBorderColor %>">
<TR bgcolor="<%= strTableBorderColor %>">
	<TD bgcolor="<% =strHeadCellColor %>" width=85%>
	<font face="<% =strDefaultFontFace %>" color="<% =strHeadFontColor %>" size="<% =strFooterFontSize %>"><B>已设定回复邮件提醒的文章</B>
<%
		if rPage > 1 then
			response.write "&nbsp;&nbsp;&nbsp;<a href=""cancel_replies.asp?rPage=1"" alt=""第一页"">[第一页]&nbsp;&nbsp;&nbsp;</a>"
			response.write "<a href=""cancel_replies.asp?rPage=" & rPage - 1 & """ alt=""上一页"">"
			response.write "前一页</a>&nbsp;|&nbsp;"
		else
			response.write "前一页&nbsp;|&nbsp;"
		end if
		if rPage < rPageCount then 
			response.write "<a href=""cancel_replies.asp?rPage=" & rPage + 1 & """ alt=""下一页"">"
			response.write "下一页</a>"
			response.write "&nbsp;&nbsp;&nbsp;<a href=""cancel_replies.asp?rPage=" & rPageCount & """ alt=""最后一页"">[最后一页]</a>"
		else
			response.write "下一页"
		end if
		response.write "&nbsp;&nbsp;&nbsp;[页数：" & rPage & "/" & rPageCount & "]"
							
					%>
	</td>
	<td bgcolor="<% =strHeadCellColor %>"><font face="<% =strDefaultFontFace %>" color="<% =strHeadFontColor %>" size="<% =strFooterFontSize %>"><a href="cancel_replies.asp?mode=deleteAllReply">全部删除</a></td>
</TR>
<%
	If rsReply.Eof OR rsReply.Bof Then
		Response.Write "<tr><td colspan=2 bgcolor='" & strForumCellColor & "'><font face='" & DefaultFontFace & "' size='2'><b>暂时没有任何回复</b></font></td></tr>"
		boolNoBookmarks = TRUE
	Else
		rsReply.AbsolutePage = rPage
		boolNoBookmarks = FALSE
		' List bookmarks
		Dim i, CColor, TxtMessage
		i=0
			rec = 1
			do until rsReply.eof or rec = (bkReplyPageSize +1)
			if i = 0 then
				CColor = strAltForumCellColor
			else
				CColor = strForumCellColor
			End if
			TxtMessage = ChkString(rsReply("R_Message"),"display")
			Response.Write "<td bgcolor='" & CColor & "'><font face='" & strDefaultFontFace & "' color=" & strForumFontColor & ">" 
			Response.Write "<a href='link.asp?topic_id=" & rsReply("Topic_ID") & "'>" & TxtMessage & "</a>"
			Response.Write "<br>&nbsp;作者： " & getMemberName(rsReply("R_AUTHOR")) & " 发表时间： " & chkDate(rsReply("R_Date")) & "</td>"
			Response.Write "<TD bgcolor=" & CColor & " align=center><font face=" & strDefaultFontFace & " color=" & strForumFontColor & "><input type=checkbox name=""delRequest"" value=""" & rsReply("REPLY_ID") & """></td></TR>"
	    rec = rec + 1
		rsReply.MoveNext
	    i = i + 1
	    if i = 2 then i = 0
	   loop
	End If
%>
</TR>
</TABLE>
<div align="center"><br><input type=submit name="del" value="删除所选文章"></div>
</FORM>
<p></p>
<%'######################################### TOPICS ######################
	set rsTopic = server.CreateObject("ADODB.RecordSet")
	
	strSql = "SELECT * FROM "& strTablePrefix & "TOPICS "
	strSql = strSql & "WHERE "& strTablePrefix & "TOPICS.T_AUTHOR=" & getmemberID(strDBNTUserName)
	strSQL = strSQL & " AND T_MAIL=1"
	rsTopic.Open strSQL, my_Conn, 3
	tPage = CLng(request("tPage"))
	rsTopic.PageSize = bkTopicPageSize
	tPageCount = cInt(rsTopic.PageCount)
	if tPageCount = 0 then tPageCount = 1
	if tPage < 1 then
		tPage = 1
	end if
%>
<br>
<form Action="cancel_replies.asp?mode=deleteTopic" method=post id=form1 name=form1>
<TABLE align="center" border="0" cellPadding=3 cellSpacing=1 width=85% bgcolor="<%= strTableBorderColor %>">
<TR bgcolor="<%= strTableBorderColor %>">
	<TD bgcolor="<% =strHeadCellColor %>" width=85%>
	<font face="<% =strDefaultFontFace %>" color="<% =strHeadFontColor %>" size="<% =strFooterFontSize %>"><B>已设定回复提醒的主题</B>
<%
		if tPage > 1 then
			response.write "&nbsp;&nbsp;&nbsp;<a href=""cancel_replies.asp?tPage=1"" alt=""第一页"">[第一页]&nbsp;&nbsp;&nbsp;</a>"
			response.write "<a href=""cancel_replies.asp?tPage=" & tPage - 1 & """ alt=""上一页"">"
			response.write "上一页</a>&nbsp;|&nbsp;"
		else
			response.write "上一页&nbsp;|&nbsp;"
		end if
		if tPage < tPageCount then 
			response.write "<a href=""cancel_replies.asp?tPage=" & tPage + 1 & """ alt=""下一页"">"
			response.write "下一页</a>"
			response.write "&nbsp;&nbsp;&nbsp;<a href=""cancel_replies.asp?tPage=" & tPageCount & """ alt=""最后一页"">[最后一页]</a>"
		else
			response.write "下一页"
		end if
		response.write "&nbsp;&nbsp;&nbsp;[页数：" & tPage & "/" & tPageCount & "]"
							
					%>
	</td>
	<td bgcolor="<% =strHeadCellColor %>"><font face="<% =strDefaultFontFace %>" color="<% =strHeadFontColor %>" size="<% =strFooterFontSize %>"><a href="cancel_replies.asp?mode=deleteAllTopic">全部删除</a></td>
</TR>
<%
	If rsTopic.Eof OR rsTopic.Bof Then
		Response.Write "<tr><td colspan=2 bgcolor='" & strForumCellColor & "'><font face='" & DefaultFontFace & "' size='2'><b>暂时没有任何主题</b></font></td></tr>"
		boolNoBookmarks = TRUE
	Else
		rsTopic.AbsolutePage = tPage
		boolNoBookmarks = FALSE
		' List Topics
		i=0
			rec = 1
			do until rsTopic.eof or rec = (bkReplyPageSize +1)
			if i = 0 then
				CColor = strAltForumCellColor
			else
				CColor = strForumCellColor
			End if
			Response.Write "<td bgcolor='" & CColor & "'><font face='" & strDefaultFontFace & "' color=" & strForumFontColor & ">" 
			Response.Write "<br><a href='link.asp?topic_id=" & rsTopic("Topic_ID") & "'>" & left(rsTopic("T_SUBJECT"), 50) & "</a>"
			Response.Write "<br>&nbsp;作者：" & getMemberName(rsTopic("T_AUTHOR")) & " 发表时间： " & chkDate(rsTopic("T_Date")) & "</td>"
			Response.Write "<TD bgcolor=" & CColor & " align=center><font face=" & strDefaultFontFace & " color=" & strForumFontColor & "><input type=checkbox name=""delRequest"" value=""" & rsTopic("TOPIC_ID") & """></td></TR>"
	    rec = rec + 1
		rsTopic.MoveNext
	    i = i + 1
	    if i = 2 then i = 0
	   loop
	End If
%>
</TR>
</TABLE>
<div align="center"><br><input type=submit name="del" value="删除所选主题"></div>
</FORM>
<P>
<%set rs = nothing
end select%>