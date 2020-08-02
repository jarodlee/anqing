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
<% 
'Create New Bookmarks tableset 
'my_conn= Server.CreateObject("ADODB.Connection")
'set rs = Server.CreateObject("ADODB.Recordset")
'my_Conn.Open ConnString
'on error resume next
'StrSql = "CREATE TABLE FORUM_Bookmarks (Bookmark_ID COUNTER, "
'StrSql = StrSql & "B_MemberID REAL, B_TopicID REAL, "
'set rs = my_conn.Execute (StrSql)
'set rs = nothing
'set my_conn = nothing
'rs.Close
'my_conn.Close 

'Now here is the fun bit...

'Just add a link to the top.inc page probably near the search links and stuff:
'<a href="bookmark.asp" title="查看你的收藏夹">你的收藏夹</a> 
'Now add a link to the topic.asp page somewhere near the top:
'<a href="bookmark.asp?mode=add&id=<%=rs("Topic_ID")% >">将本主题加入收藏夹</a> 
'Then create a file called ... yes you can guess it ... bookmark.asp and put all this code in it:

 %>
<!--#INCLUDE FILE="config.asp" -->
<!--#INCLUDE FILE="inc_functions.asp" -->
<%
nRefreshTime = Request.Cookies(strCookieURL & "Reload")

if Request.form("cookie") = "1" then
        if strSetCookieToForum = 1 then	
      Response.Cookies(strCookieURL & "Reload").Path = strCookieURL
	end if
	Response.Cookies(strCookieURL & "Reload") = Request.Form("RefreshTime")
	Response.Cookies(strCookieURL & "Reload").expires = strForumTimeAdjust + 365
	nRefreshTime = Request.Form("RefreshTime")
end if

if nRefreshTime = "" then
	nRefreshTime = 0
end if
ActiveSince = Request.Cookies(strCookieURL & "ActiveSince")
'## Do Cookie stuffs with show last date
if Request.form("cookie") = "2" then
	ActiveSince = Request.Form("ShowSinceDateTime")
    if strSetCookieToForum = 1 then	
      Response.Cookies(strCookieURL & "ActiveSince").Path = strCookieURL
	end if
	Response.Cookies(strCookieURL & "ActiveSince") = ActiveSince
end if
Select Case ActiveSince
	Case "LastVisit" 
		lastDate = ""
	Case "LastHour" 
		lastDate = DateToStr(DateAdd("h",-1,strForumTimeAdjust))
	Case "LastDay" 
		lastDate = DateToStr(DateAdd("d",-1,strForumTimeAdjust))
	Case "LastWeek" 
		lastDate = DateToStr(DateAdd("w",-7,strForumTimeAdjust))
	Case "LastMonth" 
		lastDate = DateToStr(DateAdd("m",-1,strForumTimeAdjust))
	Case "Registration"
		lastDate = "All"
	Case Else
		lastDate = ""
End Select
%>
<!--#INCLUDE FILE="inc_top.asp" -->
<script language="JavaScript">
<!--
function autoReload()
{
	document.ReloadFrm.submit()
}
//-->
</script>
<script language="JavaScript">
<!--
function SetLastDate()
{
	document.LastDateFrm.submit()
}
//-->
</script>

<%
if IsEmpty(Session(strCookieURL & "last_here_date")) then
	Session(strCookieURL & "last_here_date") = ReadLastHereDate(Request.Cookies(strUniqueID & "User")("Name"))
end if
if lastDate = "" then
	lastDate = Session(strCookieURL & "last_here_date")
end if
if Request.Form("AllRead") = "Y" then
	Session(strCookieURL & "last_here_date") = ReadLastHereDate(Request.Cookies(strUniqueID & "User")("Name"))
	Session(strCookieURL & "last_here_date") = ReadLastHereDate(Request.Cookies(strUniqueID & "User")("Name"))
	lastDate = Session(strCookieURL & "last_here_date")
end if

	bkMarkPageSize = 2
	set rs = server.CreateObject("ADODB.RecordSet")
	set rsUser = server.CreateObject("ADODB.RecordSet")


If trim(strDBNTUserName)="" Then
	' No user/password so ask user
	response.write "<br><br><P><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><center><b>建立或查看收藏夹前请先登陆论坛</b><br><br>如果你尚未注册，请<a href=""policy.asp"">点击这里注册</a>"
	response.write "</P>"
	Response.End
End If

' Got password details so do it
Select Case Request.QueryString("mode")

Case "add"
	strTopic = Request.QueryString("id")
	testSql ="SELECT * FROM "& strTablePrefix & "Bookmarks WHERE B_MemberID=" & getmemberID(strDBNTUserName) & " and B_TopicID=" & strTopic
	rs.Open testSql, my_Conn
	If rs.BOF or rs.EOF Then

		' Bookmark doesn't already exist so add it
		insSql = "INSERT INTO "& strTablePrefix & "Bookmarks (B_MemberID, B_TopicID) VALUES ("
		insSql = insSql & getmemberID(strDBNTUserName) & ", "
		insSql = insSql & strTopic & ")"
		
		my_conn.Execute (insSql)
	End If
	%>
	<br><center>
	<meta http-equiv="Refresh" content="0; URL=<% response.write "link.asp?topic_id=" & strTopic %>">
	<center>
	
	<p><font size="<% =strDefaultFontSize %>"><b>已将主题加入收藏夹！</b></font></p>
	<p><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">你可以浏览任何你加入收藏夹的主题<br>在任何的时间，只要点选"我的收藏夹"。<br>现在系统把你送回论坛或主题，如果没有反应请<a href="<% response.write "link.asp?topic_id=" & strTopic %>">点击这里</a>.
	</font></p></center>
	<div align="center"><font face="<% =strDefaultFontFace %>" size="2"><a href="default.asp">返回论坛</a></font>
</div>
	<P>&nbsp;<p>&nbsp;<p>
	<%

Case "delete"
	delBookmark = split(Request.Form("delBookmark"), ",")
	for i = 0 to ubound(delBookmark)
		' Delete selected topic bookmarks
		delSQL = "Delete From "& strTablePrefix & "Bookmarks where Bookmark_ID = " & cint(delBookmark(i))
    my_conn.Execute delSQL
	next
	response.write "<br><center><font face=""" & strDefaultFontFace & """ size=3><b>选择的收藏夹已删除！</b></font><br>"
	response.write "<a href=""default.asp"">返回论坛首页</a></center>"
Case "deleteAll"
	delBookmark = split(Request.Form("delBookmark"), ",")
		delSQL = "Delete From "& strTablePrefix & "Bookmarks where B_MemberID = " & getmemberID(strDBNTUserName)
    my_conn.Execute delSQL

	response.write "<br><center><font face=""" & strDefaultFontFace & """ size=3><b>所有的收藏都被删除了！</b></font><br>"
	response.write "<a href=""default.asp"">返回论坛首页</a></center>"

Case default
%>
<%
		
	If trim(strDBNTUserName)="" Then%>
	<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">错误的会员名或密码</font></p>
	<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript:history.go(-1)">返回重试</a></font></p>
<%	Response.end
	end if
	strSql = "SELECT "& strTablePrefix & "Bookmarks.Bookmark_ID, "& strTablePrefix & "Category.Cat_Name, "& strTablePrefix & "Forum.F_SUBJECT, "& strTablePrefix & "Topics.Topic_ID, "& strTablePrefix & "Topics.T_subject, "& strTablePrefix & "Topics.T_Replies, "& strTablePrefix & "Topics.T_date, "& strTablePrefix & "Topics.T_Last_Post,"& strMemberTablePrefix & "MEMBERS.M_NAME "
	strSQL = strSQL & "," & strTablePrefix & "TOPICS.FORUM_ID, " & strTablePrefix & "TOPICS.CAT_ID, " & strTablePrefix & "TOPICS.T_LAST_POST_AUTHOR "
	strSql = strSql & "FROM ("& strTablePrefix & "Category INNER JOIN "& strTablePrefix & "Forum ON "& strTablePrefix & "Category.Cat_ID = "& strTablePrefix & "Forum.CAT_ID) INNER JOIN ("& strMemberTablePrefix & "Members INNER JOIN ("& strTablePrefix & "Bookmarks INNER JOIN "& strTablePrefix & "Topics ON "& strTablePrefix & "Bookmarks.B_TopicID = "& strTablePrefix & "Topics.Topic_ID) ON "
	strSql = strSql & strMemberTablePrefix & "Members.Member_id = "& strTablePrefix & "Topics.T_AUTHOR) ON "& strTablePrefix & "Forum.Forum_ID = "& strTablePrefix & "Topics.Forum_id "
	strSql = strSql & "WHERE ("& strTablePrefix & "Bookmarks.B_MemberID=" & getmemberID(strDBNTUserName) & ")"
	if lastDate <> "All" then strsql = strsql & " AND " & strTablePrefix & "TOPICS.T_LAST_POST > '" & lastDate & "'"
	rs.Open strSQL, my_Conn, 3

	iPage = CLng(request("iPage"))
	rs.PageSize = bkMarkPageSize

	iPageCount = cInt(rs.PageCount)

	if iPage < 1 then
		iPage = 1
	end if
			
	
%>
<table width="95%" border=0 >
  <tr>
    <td ><form name="LastDateFrm" action="bookmark.asp" method="post">

	<select name="ShowSinceDateTime" size="1" onchange="SetLastDate();">
        <option value="LastVisit" <% if ActiveSince = "LastVisit" or ActiveSince = "" then Response.Write(" SELECTED")%>>&nbsp;自从上次光临时间：<%= ChkDate(Session(strCookieURL & "last_here_date")) %> <% =ChkTime(Session(strCookieURL & "last_here_date")) %>&nbsp;</option>
        <option value="LastHour"  <% if ActiveSince = "LastHour" then Response.Write(" SELECTED")%>>&nbsp;自从一个小时前</option>
        <option value="LastDay"   <% if ActiveSince = "LastDay" then Response.Write(" SELECTED")%>>&nbsp;自从昨天</option>
        <option value="LastWeek"  <% if ActiveSince = "LastWeek" then Response.Write(" SELECTED")%>>&nbsp;自从上星期</option>
        <option value="LastMonth" <% if ActiveSince = "LastMonth" then Response.Write(" SELECTED")%>>&nbsp;自从上个月<//option>
		<option value="Registration" <% if ActiveSince = "Registration" or ActiveSince = "All" then Response.Write(" SELECTED")%>>&nbsp;自从注册</option>
     </select>
    <input type="hidden" name="Cookie" value="2">
    </form>
</font>
    </td>

    <td align="right">
    <form name="ReloadFrm" action="bookmark.asp" method="post"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"> 
    <select name="RefreshTime" size="1" onchange="autoReload();">
        <option value="0"  <% if nRefreshTime = "0" then Response.Write(" SELECTED")%>>不要自动重新整理</option>
        <option value="1"  <% if nRefreshTime = "1" then Response.Write(" SELECTED")%>>每分钟都重新整理</option>
        <option value="5"  <% if nRefreshTime = "5" then Response.Write(" SELECTED")%>>每隔五分钟重新整理</option>
        <option value="10" <% if nRefreshTime = "10" then Response.Write(" SELECTED")%>>每隔十分钟重新整理</option>
        <option value="15" <% if nRefreshTime = "15" then Response.Write(" SELECTED")%>>每隔十五分钟重新整理</option>
        <option value="30" <% if nRefreshTime = "30" then Response.Write(" SELECTED")%>>每隔半个钟头重新整理</option>
    </select>
    <input type="hidden" name="Cookie" value="1">
    </font>
    </form>
	
    </td>
  </tr>

</table>

<br>
<form Action="bookmark.asp?mode=delete" method=post id=form1 name=form1>
<TABLE align="center" border="0" cellPadding=3 cellSpacing=1 width=95% bgcolor="<%= strTableBorderColor %>">
<TR bgcolor="<%= strTableBorderColor %>">
	<TD bgcolor="<% =strHeadCellColor %>" width=85%>
	<font face="<% =strDefaultFontFace %>" color="<% =strHeadFontColor %>" size="<% =strFooterFontSize %>"><B>主题</B>
						<%
							if iPage > 1 then
								response.write "&nbsp;&nbsp;&nbsp;<a href=""bookmark.asp?iPage=1"" alt=""第一页"">[第一页]&nbsp;&nbsp;&nbsp;</a>"
								response.write "<a href=""bookmark.asp?iPage=" & iPage - 1 & """ alt=""上一页"">"
								response.write "←</a>&nbsp;|&nbsp;"
							else
								response.write "←&nbsp;|&nbsp;"
							end if
							if iPage < iPageCount then 
								response.write "<a href=""bookmark.asp?iPage=" & iPage + 1 & """ alt=""下一页"">"
								response.write "→ </a>"
								response.write "&nbsp;&nbsp;&nbsp;<a href=""bookmark.asp?iPage=" & iPageCount & """ alt=""最后一页"">[最后一页]</a>"
							else
								response.write "→ "
							end if
							response.write "&nbsp;&nbsp;&nbsp;[页:" & iPage & "/" & iPageCount & "]"
							
					%>
	</td>
	<td bgcolor="<% =strHeadCellColor %>" align="center"><font face="<% =strDefaultFontFace %>" color="<% =strHeadFontColor %>" size="<% =strFooterFontSize %>"><a href="bookmark.asp?mode=deleteAll" title="删除全部收藏">全删</a></td>
</TR>

<%
	If rs.Eof OR rs.Bof Then
		Response.Write "<tr><td colspan=2 bgcolor='" & strForumCellColor & "'><font face='" & DefaultFontFace & "' size='2'>目前尚无任何收藏</font></td></tr>"
		boolNoBookmarks = TRUE
	Else
		rs.AbsolutePage = iPage
		boolNoBookmarks = FALSE
		' List bookmarks
		Dim i, CCsolor
		i=0
			rec = 1
			do until rs.eof or rec = (bkMarkPageSize +1)
			if i = 0 then
				CColor = strAltForumCellColor
			else
				CColor = strForumCellColor
			End if
			
			Response.Write "<td bgcolor='" & CColor & "'><font face='" & strDefaultFontFace & "' color=" & strForumFontColor & " size='2'>" & rs("Cat_Name") & " - " & rs("F_SUBJECT") & " 的： " 
			Response.Write "<a href='link.asp?topic_id=" & rs("Topic_ID") & "'>" & left(rs("T_Subject"), 50) & "</a><font size=-2>"

			If rs("T_Replies") > strPageSize Then
			' More than X replies on the page and therefore must be more than one page
				Response.Write "&nbsp;&nbsp;-&nbsp;&nbsp;Jump&nbsp;to&nbsp;page:&nbsp;"
				Dim strRef, nCounter, nPages
				nPages = fix(rs("T_Replies") / strPageSize)
				If ( ( nPages * strPageSize ) <> rs("T_Replies") ) Then
					nPages  = nPages + 1
				End If
				For nCounter = 1 to nPages
					strRef="<a href='topic.asp"
					strRef=strRef & "?whichpage=" & nCounter
					strRef=strRef & "&topic_id=" & rs("Topic_ID")
					strRef=strRef & "&forum_id=" & rs("Forum_ID")
					strRef=strRef & "&cat_id=" & rs("CAT_ID")
					strRef=strRef & "'>" & nCounter & "</a>"
					Response.Write strRef & "&nbsp;"
				Next
			End If

			Response.Write "<br>此文章由 " & getMemberName(rs("T_LAST_POST_AUTHOR")) & " 发表于<b> " & chkDate(rs("T_LAST_POST")) & " </b></font></td>"
			Response.Write "<TD bgcolor=" & CColor & " align=center><font face=" & strDefaultFontFace & " color=" & strForumFontColor & " size=2><input type=checkbox name=""delBookmark"" value=""" & rs("Bookmark_ID") & """></td></TR>"
	    rec = rec + 1
		rs.MoveNext
	    i = i + 1
	    if i = 2 then i = 0
	   loop
	End If
%>
</TR>
</TABLE>
<p><div align="center">

<input type=submit name="del" value="删除所选的收藏">
</FORM>
</div></p><P>

<P>
<div align="center"><font face="<% =strDefaultFontFace %>" size="2"><a href="default.asp">返回论坛首页</a></font>
</div>
<SCRIPT>
<!--
if (document.ReloadFrm.RefreshTime.options[document.ReloadFrm.RefreshTime.selectedIndex].value > 0) {
	reloadTime = 60000 * document.ReloadFrm.RefreshTime.options[document.ReloadFrm.RefreshTime.selectedIndex].value
	self.setInterval('autoReload()', 60000 * document.ReloadFrm.RefreshTime.options[document.ReloadFrm.RefreshTime.selectedIndex].value)
}
//-->
</SCRIPT>

<%
End Select
set rs = nothing
set rsUser = nothing
%>
<!--#include file="inc_footer.asp"-->
