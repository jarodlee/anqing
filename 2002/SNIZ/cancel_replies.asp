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
	response.write "<div align=""center""><font face=""" & strDefaultFontFace & """ size=3><b>��ѡҪ���������</b></font>"
	response.write "<br><a href=""default.asp"">������̳��ҳ</a></div>"
	response.write "<meta http-equiv=""Refresh"" content=""0; URL=cancel_replies.asp"">"
	Response.End

Case "deleteAllReply"
	delSQL = "UPDATE "& strTablePrefix & "REPLY SET R_MAIL = 0 WHERE REPLY_ID IN ("
	delSQL = delSQL & "SELECT REPLY_ID FROM "& strTablePrefix & "REPLY "
	delSQL = delSQL & "WHERE "& strTablePrefix & "REPLY.R_AUTHOR=" & getmemberID(strDBNTUserName) & " AND R_MAIL=1)"
    my_conn.Execute delSQL
	response.write "<center><font face=""" & strDefaultFontFace & """ size=3><b>����Ҫ����������</b></font></center>"
	response.write "<div align=""center""><a href=""default.asp"">������̳��ҳ</a></div>"
	response.write "<meta http-equiv=""Refresh"" content=""0; URL=cancel_replies.asp"">"
	Response.End

Case "deleteTopic"
	delRequest = split(Request.Form("delRequest"), ",")
	for i = 0 to ubound(delRequest)
		delSQL = "UPDATE "& strTablePrefix & "TOPICS SET T_MAIL=0 WHERE TOPIC_ID = " & cint(delRequest(i))
	    my_conn.Execute delSQL
	next
	response.write "<center><font face=""" & strDefaultFontFace & """ size=3><b>��ѡҪ���������</b></font></center>"
	response.write "<div align=""center""><a href=""default.asp"">������̳��ҳ</a></div>"
	response.write "<meta http-equiv=""Refresh"" content=""0; URL=cancel_replies.asp"">"
	Response.End

Case "deleteAllTopic"
	delSQL = "UPDATE "& strTablePrefix & "TOPICS SET T_MAIL = 0 WHERE TOPIC_ID IN ("
	delSQL = delSQL & "SELECT TOPIC_ID FROM "& strTablePrefix & "TOPICS "
	delSQL = delSQL & "WHERE "& strTablePrefix & "TOPICS.T_AUTHOR=" & getmemberID(strDBNTUserName) & " AND T_MAIL=1)"
    my_conn.Execute delSQL
	response.write "<center><font face=""" & strDefaultFontFace & """ size=3><b>��ѡ���ӽ��������</b></font></center>"
	response.write "<div align=""center""><a href=""default.asp"">������̳��ҳ</a></div>"
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
	<font face="<% =strDefaultFontFace %>" color="<% =strHeadFontColor %>" size="<% =strFooterFontSize %>"><B>���趨�ظ��ʼ����ѵ�����</B>
<%
		if rPage > 1 then
			response.write "&nbsp;&nbsp;&nbsp;<a href=""cancel_replies.asp?rPage=1"" alt=""��һҳ"">[��һҳ]&nbsp;&nbsp;&nbsp;</a>"
			response.write "<a href=""cancel_replies.asp?rPage=" & rPage - 1 & """ alt=""��һҳ"">"
			response.write "ǰһҳ</a>&nbsp;|&nbsp;"
		else
			response.write "ǰһҳ&nbsp;|&nbsp;"
		end if
		if rPage < rPageCount then 
			response.write "<a href=""cancel_replies.asp?rPage=" & rPage + 1 & """ alt=""��һҳ"">"
			response.write "��һҳ</a>"
			response.write "&nbsp;&nbsp;&nbsp;<a href=""cancel_replies.asp?rPage=" & rPageCount & """ alt=""���һҳ"">[���һҳ]</a>"
		else
			response.write "��һҳ"
		end if
		response.write "&nbsp;&nbsp;&nbsp;[ҳ����" & rPage & "/" & rPageCount & "]"
							
					%>
	</td>
	<td bgcolor="<% =strHeadCellColor %>"><font face="<% =strDefaultFontFace %>" color="<% =strHeadFontColor %>" size="<% =strFooterFontSize %>"><a href="cancel_replies.asp?mode=deleteAllReply">ȫ��ɾ��</a></td>
</TR>
<%
	If rsReply.Eof OR rsReply.Bof Then
		Response.Write "<tr><td colspan=2 bgcolor='" & strForumCellColor & "'><font face='" & DefaultFontFace & "' size='2'><b>��ʱû���κλظ�</b></font></td></tr>"
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
			Response.Write "<br>&nbsp;���ߣ� " & getMemberName(rsReply("R_AUTHOR")) & " ����ʱ�䣺 " & chkDate(rsReply("R_Date")) & "</td>"
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
<div align="center"><br><input type=submit name="del" value="ɾ����ѡ����"></div>
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
	<font face="<% =strDefaultFontFace %>" color="<% =strHeadFontColor %>" size="<% =strFooterFontSize %>"><B>���趨�ظ����ѵ�����</B>
<%
		if tPage > 1 then
			response.write "&nbsp;&nbsp;&nbsp;<a href=""cancel_replies.asp?tPage=1"" alt=""��һҳ"">[��һҳ]&nbsp;&nbsp;&nbsp;</a>"
			response.write "<a href=""cancel_replies.asp?tPage=" & tPage - 1 & """ alt=""��һҳ"">"
			response.write "��һҳ</a>&nbsp;|&nbsp;"
		else
			response.write "��һҳ&nbsp;|&nbsp;"
		end if
		if tPage < tPageCount then 
			response.write "<a href=""cancel_replies.asp?tPage=" & tPage + 1 & """ alt=""��һҳ"">"
			response.write "��һҳ</a>"
			response.write "&nbsp;&nbsp;&nbsp;<a href=""cancel_replies.asp?tPage=" & tPageCount & """ alt=""���һҳ"">[���һҳ]</a>"
		else
			response.write "��һҳ"
		end if
		response.write "&nbsp;&nbsp;&nbsp;[ҳ����" & tPage & "/" & tPageCount & "]"
							
					%>
	</td>
	<td bgcolor="<% =strHeadCellColor %>"><font face="<% =strDefaultFontFace %>" color="<% =strHeadFontColor %>" size="<% =strFooterFontSize %>"><a href="cancel_replies.asp?mode=deleteAllTopic">ȫ��ɾ��</a></td>
</TR>
<%
	If rsTopic.Eof OR rsTopic.Bof Then
		Response.Write "<tr><td colspan=2 bgcolor='" & strForumCellColor & "'><font face='" & DefaultFontFace & "' size='2'><b>��ʱû���κ�����</b></font></td></tr>"
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
			Response.Write "<br>&nbsp;���ߣ�" & getMemberName(rsTopic("T_AUTHOR")) & " ����ʱ�䣺 " & chkDate(rsTopic("T_Date")) & "</td>"
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
<div align="center"><br><input type=submit name="del" value="ɾ����ѡ����"></div>
</FORM>
<P>
<%set rs = nothing
end select%>