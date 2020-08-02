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
<%
'## Do Cookie stuffs with reload
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

WhichTopics = Request.Form("ShowWhichTopics")
if WhichTopics = "" then WhichTopics = "All"

Select Case ActiveSince
	Case "LastVisit" 
		lastDate = ""
	Case "LastHour" 
		lastDate = DateToStr(DateAdd("h",-1,strForumTimeAdjust))
	Case "LastDay" 
		lastDate = DateToStr(DateAdd("d",-1,strForumTimeAdjust))
	Case "LastWeek" 
		lastDate = DateToStr(DateAdd("ww",-7,strForumTimeAdjust))
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
<script language="JavaScript">
<!--
function jumpTo(s) {if (s.selectedIndex != 0) top.location.href = s.options[s.selectedIndex].value;return 1;}
// -->
</script>
<%
if IsEmpty(Session(strCookieURL & "last_here_date")) then
	Session(strCookieURL & "last_here_date") = ReadLastHereDate(strDBNTUserName)
end if
if lastDate = "" then
	lastDate = Session(strCookieURL & "last_here_date")
end if
if Request.Form("AllRead") = "Y" then
	Session(strCookieURL & "last_here_date") = ReadLastHereDate(Request.Cookies(strUniqueID & "User")("Name"))
	Session(strCookieURL & "last_here_date") = ReadLastHereDate(Request.Cookies(strUniqueID & "User")("Name"))
	lastDate = Session(strCookieURL & "last_here_date")
	ActiveSince = " "
end if

'## Forum_SQL - Get all active topics from last visit
					strsql = "SELECT " & strTablePrefix & "FORUM.FORUM_ID, " & strTablePrefix & "FORUM.F_SUBJECT, " & strTablePrefix & "FORUM.CAT_ID, " & strTablePrefix & "TOPICS.TOPIC_ID, " 
					strsql = strsql & strTablePrefix & "TOPICS.T_SUBJECT, " & strTablePrefix & "TOPICS.T_STATUS,  " & strTablePrefix & "TOPICS.T_LAST_POST,  " & strTablePrefix & "TOPICS.T_AUTHOR,  "
					strSql = strSql & strTablePrefix & "TOPICS.T_REPLIES, " & strTablePrefix & "TOPICS.T_LAST_POST_AUTHOR, " & strTablePrefix & "TOPICS.T_VIEW_COUNT  "
					strsql = strsql & " FROM ((" & strTablePrefix & "FORUM LEFT JOIN " & strTablePrefix & "TOPICS "
					strsql = strsql & " ON " & strTablePrefix & "FORUM.FORUM_ID = " & strTablePrefix & "TOPICS.FORUM_ID) LEFT JOIN " & strTablePrefix & "REPLY "
					strsql = strsql & " ON " & strTablePrefix & "TOPICS.TOPIC_ID = " & strTablePrefix & "REPLY.TOPIC_ID) "
					if WhichTopics = "My" then
						strsql = strsql & " WHERE  " 
						strsql = strsql & " (" & strTablePrefix & "TOPICS.T_AUTHOR = " & getmemberID(strDBNTUserName) & " "
						strsql = strsql & " OR " & strTablePrefix & "REPLY.R_AUTHOR = " & getmemberID(strDBNTUserName) & ")"
						if lastdate <> "All" then 	strsql = strsql & " AND " & strTablePrefix & "TOPICS.T_LAST_POST > '" & lastDate & "'"
					elseif WhichTopics = "All" then
						strSql = strSql & "WHERE " & strTablePrefix & "TOPICS.T_LAST_POST > '" & lastDate & "'"
					elseif WhichTopics = "Bookmark" then
						strsql = strsql & " WHERE  " 
						strsql = strsql & " EXISTS (SELECT * FROM " & strTablePrefix & "BOOKMARKS WHERE " & strTablePrefix & "BOOKMARKS.B_MemberID =" & getmemberID(strDBNTUserName) & " "
						strsql = strsql & " AND " & strTablePrefix & "BOOKMARKS.B_TopicID = " & strTablePrefix & "TOPICS.TOPIC_ID )"
						if lastdate <> "All" then 	strsql = strsql & " AND " & strTablePrefix & "TOPICS.T_LAST_POST > '" & lastDate & "'"
					end if
					strSql = strSql & " ORDER BY " & strTablePrefix & "TOPICS.FORUM_ID, " & strTablePrefix & "TOPICS.T_LAST_POST DESC;"


set rs = my_Conn.Execute (strSql)

%>
<table width="100%" border=0>
  <tr>
    <td ><form name="LastDateFrm" action="active.asp" method="post">
	<select name="ShowWhichTopics" size="1" onchange="SetLastDate();"> 
        <option value="All" <% if WhichTopics = "All" or WhichTopics = "" then Response.Write(" SELECTED")%>>&nbsp;全 部</option>
	<% If strDBNTUserNAme <> "" then %>
        <option value="My"  <% if WhichTopics = "My" then Response.Write(" SELECTED")%>>&nbsp;我的</option>
		<% on error resume next
			my_conn.execute "SELECT * FROM " & strTablePrefix & "BOOKMARKS" 
			if err.number = 0 then%>
			<option value="Bookmark"  <% if WhichTopics = "Bookmark" then Response.Write(" SELECTED")%>>&nbsp;收藏夹中的</option>
			<% Else
			'do nothing
			 End If 
		End if%>
     </select>
	<select name="ShowSinceDateTime" size="1" onchange="SetLastDate();">
        <option value="LastVisit" <% if ActiveSince = "LastVisit" or ActiveSince = "" then Response.Write(" SELECTED")%>>&nbsp;自从上次 <%= ChkDate(Session(strCookieURL & "last_here_date")) %> 光临之后&nbsp;</option>
        <option value="LastHour"  <% if ActiveSince = "LastHour" then Response.Write(" SELECTED")%>>&nbsp;自从前一个钟头</option>
        <option value="LastDay"   <% if ActiveSince = "LastDay" then Response.Write(" SELECTED")%>>&nbsp;自从昨天</option>
        <option value="LastWeek"  <% if ActiveSince = "LastWeek" then Response.Write(" SELECTED")%>>&nbsp;自从上个星期</option>
        <option value="LastMonth" <% if ActiveSince = "LastMonth" then Response.Write(" SELECTED")%>>&nbsp;自从上个月<//option>
		<% If WhichTopics <> "All" then %>
				<option value="Registration" <% if ActiveSince = "Registration" or ActiveSince = "All" then Response.Write(" SELECTED")%>>&nbsp;自从注册</option>
		<% End If %>
     </select>
    <input type="hidden" name="Cookie" value="2">
    </form>
</font>
    </td>

    <td align="right">
    <form name="ReloadFrm" action="active.asp" method="post"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"> 
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
<table border="0" width="100%" cellspacing="1" cellpadding="2" align="center" bgcolor="<% =strTableBorderColor %>">
      <tr>
       <form name="MarkRead" action="active.asp" method="post">
		<input type="hidden" name="AllRead" value="Y">
		<td align="center" bgcolor="<% =strHeadCellColor %>" valign="center" width="20">
	   <%If not(rs.EOF or rs.BOF) and (mLev > 0) then %>
		<input type="image" src="<%=strImageURL %>icon_topic_all_read.gif" value="Mark all read" id=submit1 name=submit1 alt="Mark all topics as read" height="15" width="15" border="0" hspace="0" onclick>
        <% else %>
			&nbsp;
        <% end if %>
		</td>			
        </form>
        <td align="center" bgcolor="<% =strHeadCellColor %>" width="50%"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">主题</font></td>
        <td align="center" bgcolor="<% =strHeadCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">作者</font></td>
        <td align="center" bgcolor="<% =strHeadCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">回复</font></td>
        <td align="center" bgcolor="<% =strHeadCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">人气</font></td>
        <td align="center" bgcolor="<% =strHeadCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">最新文章</font></td>
<% if (mlev = 4 or mlev = 3) or (lcase(strNoCookies) = "1") then %>
        <td align="center" bgcolor="<% =strHeadCellColor %>"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">&nbsp;</font></b></td>
<% end if %>
      </tr>
<%If rs.EOF or rs.BOF then %>
      <tr>
        <td colspan="6" bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>没有找到任何有新文章的主题</b></font></td>
<% if (mlev = 4 or mlev = 3) or (lcase(strNoCookies) = "1") then %>
        <td align="center" bgcolor="<% =strForumCellColor %>"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">&nbsp;</font></b></td>
<% end if %>
      </tr>
<%else
	currForum = 0 
	fDisplayCount = 0 
	currTopic = 0
	do until rs.EOF

'## Forum_SQL - Find out if the Category is Locked or Un-Locked and if it Exists
strSql = "SELECT " & strTablePrefix & "CATEGORY.CAT_STATUS " 
strSql = strSql & " FROM " & strTablePrefix & "CATEGORY "
strSql = strSql & " WHERE " & strTablePrefix & "CATEGORY.CAT_ID = " & rs("CAT_ID")

set rsCStatus = my_Conn.Execute (StrSql)

'## Forum_SQL - Find out if the Topic is Locked or Un-Locked and if it Exists
strSql = "SELECT " & strTablePrefix & "FORUM.F_STATUS " 
strSql = strSql & " FROM " & strTablePrefix & "FORUM "
strSql = strSql & " WHERE " & strTablePrefix & "FORUM.FORUM_ID = " & rs("FORUM_ID")
if currTopic <> rs("TOPIC_ID") then
set rsFStatus = my_Conn.Execute (StrSql)
		if ChkForumAccess(rs("FORUM_ID")) then
			if (mLev = 4) or ((chkForumModerator(rs("FORUM_ID"), Session(strCookieURL & "username"))= "1") and mLev = 3) or ((chkForumModerator(rs("FORUM_ID"), Request.Cookies(strUniqueID & "User")("Name")) = "1") and mlev = 3) or (lcase(strNoCookies) = "1") then
 				AdminAllowed = 1
 			else   
 				AdminAllowed = 0
 			end if
			fDisplayCount = fDisplayCount + 1
			if currForum <> rs("FORUM_ID") then %>
				<tr>
					<td height="20" colspan="6" bgcolor="<% =strCategoryCellColor %>" valign="center" ><a href="<% Response.Write("FORUM.asp?FORUM_ID=" & rs("FORUM_ID") & "&CAT_ID=" & rs("CAT_ID") & "&Forum_Title=" & ChkString(rs("F_SUBJECT"),"urlpath")) %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>"><b><% =ChkString(rs("F_SUBJECT"),"display") %></b></font></a></td>
<%					if (AdminAllowed = 1) or (lcase(strNoCookies) = "1") then %>
						<td align="center" bgcolor="<% =strCategoryCellColor %>" nowrap valign="center"><% ForumAdminOptions() %></td>
					<%elseif (mLev = 3) then%>
						<td align="center" bgcolor="<% =strCategoryCellColor %>" nowrap valign="center">&nbsp;</td>
<%					end if %>
				</tr>
<%			end if %>
			<tr>
			<td bgcolor="<% =strForumCellColor %>" align="center"  valign="center">
<%			if rsCStatus("CAT_STATUS") <> 0 and rsFstatus("F_STATUS") <> 0 and rs("T_STATUS") <> 0 then
				if lcase(strHotTopic) = "1" then
					if rs("T_REPLIES") >= intHotTopicNum Then %>
						<a href="topic.asp?TOPIC_ID=<% =rs("TOPIC_ID") %>&FORUM_ID=<% =rs("FORUM_ID") %>&CAT_ID=<% =rs("CAT_ID") %>&Topic_Title=<% =ChkString(left(rs("T_SUBJECT"), 50),"urlpath") %>&Forum_Title=<% =ChkString(rs("F_SUBJECT"),"urlpath") %>"><img src="<%=strImageURL %>icon_folder_new_hot.gif" height="15" width="15" border="0" hspace="0" alt="热门主题"></a>
<%					else%>
						<a href="topic.asp?TOPIC_ID=<% =rs("TOPIC_ID") %>&FORUM_ID=<% =rs("FORUM_ID") %>&CAT_ID=<% =rs("CAT_ID") %>&Topic_Title=<% =ChkString(left(rs("T_SUBJECT"), 50),"urlpath") %>&Forum_Title=<% =ChkString(rs("F_SUBJECT"),"urlpath") %>"><img src="<%=strImageURL %>icon_folder_new.gif" alt="发表新主题" border="0"></a>
<%					end if
				end if
			else %>
			<a href="topic.asp?TOPIC_ID=<% =rs("TOPIC_ID") %>&FORUM_ID=<% =rs("FORUM_ID") %>&CAT_ID=<%=rs("CAT_ID") %>&Topic_Title=<% =ChkString(left(rs("T_SUBJECT"), 50),"urlpath") %>&Forum_Title=<%=ChkString(rs("F_SUBJECT"),"urlpath") %>"><img src="<%=strImageURL %>icon_folder_new_locked.gif"
<% 			if rsCStatus("CAT_STATUS") = 0 then 
				Response.Write ("alt='分类已锁定'")
			elseif rsFStatus("F_STATUS") = 0 then 
				Response.Write ("alt='讨论区已锁定'")
			else
				Response.Write ("alt='主题已锁定'")
			end if %>
			border="0"></a>
<%			end if %>
			</td>
			<td bgcolor="<% =strForumCellColor %>" valign="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="topic.asp?TOPIC_ID=<% =rs("TOPIC_ID") %>&FORUM_ID=<% =rs("FORUM_ID") %>&CAT_ID=<% =rs("CAT_ID") %>&Topic_Title=<% =ChkString(left(rs("T_SUBJECT"), 50),"urlpath") %>&Forum_Title=<% =ChkString(rs("F_SUBJECT"),"urlpath") %>"><% =ChkString(left(rs("T_SUBJECT"), 50),"display") %></a>&nbsp;</font><% if strShowPaging = "1" then TopicPaging() end if%></td>
			<td bgcolor="<% =strForumCellColor %>" valign="center" align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strForumFontColor %>"><% =ChkString(getMemberName(rs("T_AUTHOR")),"display") %></font></td>
			<td bgcolor="<% =strForumCellColor %>" valign="center" align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strForumFontColor %>"><% =rs("T_REPLIES") %></font></td>
			<td bgcolor="<% =strForumCellColor %>" valign="center" align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strForumFontColor %>"><% =rs("T_VIEW_COUNT") %></font></td>
			<%
			if rs("T_LAST_POST_AUTHOR") = "" then
				strLastAuthor = ""
			else
				strLastAuthor = "<br>作者：" 
				if strUseExtendedProfile then
					strLastAuthor = strLastAuthor & "<a href=""pop_profile.asp?mode=display&id="& rs("T_LAST_POST_AUTHOR") & """>"
				else
					strLastAuthor = strLastAuthor & "<a href=""JavaScript:openWindow2('pop_profile.asp?mode=display&id=" & rs("T_LAST_POST_AUTHOR") & "')"">"
				end if
				strLastAuthor = strLastAuthor & getMemberName(rs("T_LAST_POST_AUTHOR")) & "</a>"
			end if
			%>
			<td bgcolor="<% =strForumCellColor %>" valign="center" align="center" nowrap><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>" color="<% =strForumFontColor %>"><b><% =ChkDate(rs("T_LAST_POST")) %></b>&nbsp;<% =ChkTime(rs("T_LAST_POST")) %><%=strLastAuthor%></font></td>
<%			if (AdminAllowed = 1) or (lcase(strNoCookies) = "1") then %>
				<td bgcolor="<% =strForumCellColor %>" valign="center" align="center" nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
<%				if rsCStatus("CAT_STATUS") = 0 then %>
					<a href="JavaScript:openWindow('pop_open.asp?mode=Category&CAT_ID=<% =rs("CAT_ID") %>')"><img src="<%=strImageURL %>icon_unlock.gif" alt="解开分类锁定" border="0" hspace="0"></a>
<%				else 
					if rsFStatus("F_STATUS") = 0 then %>
						<a href="JavaScript:openWindow('pop_open.asp?mode=Forum&FORUM_ID=<% =rs("FORUM_ID") %>&CAT_ID=<% =rs("CAT_ID") %>&Forum_Title=<% =ChkString(rs("F_SUBJECT"),"JSurlpath")%>')"><img src="<%=strImageURL %>icon_unlock.gif" alt="解开讨论区锁定" border="0" hspace="0"></a>
<%					else 
						if rs("T_STATUS") <> 0 then %>
							<a href="JavaScript:openWindow('pop_lock.asp?mode=Topic&TOPIC_ID=<% =rs("TOPIC_ID")%>&FORUM_ID=<% =rs("FORUM_ID") %>&CAT_ID=<% =rs("CAT_ID") %>&Topic_Title=<% =ChkString(rs("T_SUBJECT"),"JSurlpath")%>')"><img src="<%=strImageURL %>icon_lock.gif" alt="锁定讨论区" border="0" hspace="0"></a>
<%						else %>
							<a href="JavaScript:openWindow('pop_open.asp?mode=Topic&TOPIC_ID=<% =rs("TOPIC_ID")%>&FORUM_ID=<% =rs("FORUM_ID") %>&CAT_ID=<% =rs("CAT_ID") %>&Topic_Title=<% =ChkString(rs("T_SUBJECT"),"JSurlpath")%>')"><img src="<%=strImageURL %>icon_unlock.gif" alt="解开主题锁定" border="0" hspace="0"></a>
<%						end if 
					end if 
				end if 
				if (AdminAllowed = 1) or (rsCStatus("CAT_STATUS") <> 0 and rsFStatus("F_STATUS") <> 0 and rs("T_STATUS") <> 0) then %>
					<a href="post.asp?method=EditTopic&TOPIC_ID=<% =rs("TOPIC_ID") %>&FORUM_ID=<% =rs("FORUM_ID") %>&CAT_ID=<% =rs("CAT_ID") %>&auth=<% =rs("T_AUTHOR") %>&Forum_Title=<% =ChkString(rs("F_SUBJECT"),"urlpath") %>&Topic_Title=<% =ChkString(rs("T_SUBJECT"),"urlpath") %>"><img src="<%=strImageURL %>icon_pencil.gif" alt="编辑讯息" border="0" hspace="0"></a>
<%				end if %>
				<a href="JavaScript:openWindow('pop_delete.asp?mode=Topic&TOPIC_ID=<% =rs("TOPIC_ID") %>&FORUM_ID=<% =rs("FORUM_ID") %>&CAT_ID=<% =rs("CAT_ID") %>&Topic_Title=<% =ChkString(rs("T_SUBJECT"),"JSurlpath") %>')"><img src="<%=strImageURL %>icon_trashcan.gif" alt="删除主题" border="0" hspace="0"></a>
				<a href="post.asp?method=Reply&TOPIC_ID=<% =rs("TOPIC_ID") %>&FORUM_ID=<% =rs("FORUM_ID") %>&CAT_ID=<% =rs("CAT_ID") %>&Forum_Title=<% =ChkString(rs("F_SUBJECT"),"urlpath") %>&Topic_Title=<% =ChkString(left(rs("T_SUBJECT"), 50),"urlpath") %>"><img src="<%=strImageURL %>icon_reply_topic.gif" alt="回覆主题" height=15 width=15 border=0></a>
				</font></b>
				</td>
<%			elseif (mLev = 3) then%>
				<td bgcolor="<% =strForumCellColor %>" valign="center" align="center" nowrap>&nbsp;</td>
<%			end if %>
			</tr>	
<%		end if	
		currForum = rs("FORUM_ID") 
		currTopic = rs("TOPIC_ID")
end if %>

<%		rs.MoveNext 
	loop 
	if fDisplayCount = 0 then %>
		  <tr>
		 <td colspan="6" bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>暂时没有任何新主题</b></font></td>
<%		 if (AdminAllowed = 1) or (lcase(strNoCookies) = "1") then %>
			<td align="center" bgcolor="<% =strForumCellColor %>"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">&nbsp;</font></b></td>
<%		end if
	end if 
 end if %>
    </table>
<table width=100% border=0 align=center>
  <tr>
    <td>
    &nbsp;
    </td>
    <td align=right>
    <!--#INCLUDE FILE="inc_jump_to.asp" -->
    </td>
  </tr>
</table>
<SCRIPT>
<!--
if (document.ReloadFrm.RefreshTime.options[document.ReloadFrm.RefreshTime.selectedIndex].value > 0) {
	reloadTime = 60000 * document.ReloadFrm.RefreshTime.options[document.ReloadFrm.RefreshTime.selectedIndex].value
	self.setInterval('autoReload()', 60000 * document.ReloadFrm.RefreshTime.options[document.ReloadFrm.RefreshTime.selectedIndex].value)
}
//-->
</SCRIPT>
<!--#INCLUDE FILE="inc_footer.asp" -->
<% 
sub ForumAdminOptions() 

	if (AdminAllowed = 1) or (lcase(strNoCookies) = "1") then 
		if rsCStatus("CAT_STATUS") = 0 then 
			if mlev = 4 then %>
    <a href="JavaScript:openWindow('pop_open.asp?mode=Category&CAT_ID=<% =rs("CAT_ID") %>')"><img src="<%=strImageURL %>icon_folder_unlocked.gif" alt="解开分类锁定" height=15 width=15 border=0></a>
<%			else %>
    <img src="<%=strImageURL %>icon_folder_locked.gif" alt="锁定分类" height=15 width=15 border=0>
<%			end if 
		else 
			if rsFStatus("F_STATUS") <> 0 then %>
    <a href="JavaScript:openWindow('pop_lock.asp?mode=Forum&FORUM_ID=<% =rs("FORUM_ID") %>&CAT_ID=<% =rs("CAT_ID") %>&Forum_Title=<% =ChkString(rs("F_SUBJECT"),"JSurlpath")%>')"><img src="<%=strImageURL %>icon_folder_locked.gif" alt="锁定论坛" height=15 width=15 border=0></a>
<%			else %>
    <a href="JavaScript:openWindow('pop_open.asp?mode=Forum&FORUM_ID=<% =rs("FORUM_ID") %>&CAT_ID=<% =rs("CAT_ID") %>&Forum_Title=<% =ChkString(rs("F_SUBJECT"),"JSurlpath")%>')"><img src="<%=strImageURL %>icon_folder_unlocked.gif" alt="解开论坛锁定" height=15 width=15 border=0></a>
<%			end if 
		end if 
		if (rsCStatus("CAT_STATUS") <> 0 and rsFStatus("F_STATUS") <> 0) or (AdminAllowed = 1) then %>
          <a href="post.asp?method=EditForum&FORUM_ID=<% =rs("FORUM_ID") %>&CAT_ID=<% =rs("CAT_ID") %>&Forum_Title=<% =ChkString(rs("F_SUBJECT"),"urlpath") %>&type=0"><img src="<%=strImageURL %>icon_folder_pencil.gif" alt="编辑论坛属性" border="0" hspace="0"></a>
<%		end if %>
    <a href="JavaScript:openWindow('pop_delete.asp?mode=Forum&FORUM_ID=<% =rs("FORUM_ID") %>&CAT_ID=<% =rs("CAT_ID") %>&Forum_Title=<% =ChkString(rs("F_SUBJECT"),"JSurlpath") %>')"><img src="<%=strImageURL %>icon_folder_delete.gif" alt="删除论坛" height=15 width=15 border=0></a>
    <a href="post.asp?method=Topic&FORUM_ID=<% =rs("FORUM_ID")%>&CAT_ID=<% =rs("CAT_ID")%>&Forum_Title=<% =ChkString(rs("F_SUBJECT"),"urlpath") %>"><img src="<%=strImageURL %>icon_folder_new_topic.gif" alt="发表新主题" height=15 width=15 border=0></a>
<%	end if 

end sub


sub TopicPaging()
    mxpages = (rs("T_REPLIES") / strPageSize)
    if mxPages <> cint(mxPages) then
        mxpages = int(mxpages) + 1
    end if
    if mxpages > 1 then
		Response.Write("<table border=0 cellspacing=0 cellpadding=0><tr><td valign=""center""><font face=" & strDefaultFontFace & " size=" & strFooterFontSize & "><img src=""" & strImageURL & "icon_posticon.gif"" border=""0""></font></td>")
		for counter = 1 to mxpages
			ref = "<td align=right valign=""center"" bgcolor=" & strForumCellColor  & "><font face=" & strDefaultFontFace & " size=" & strFooterFontSize & ">" 
			if ((mxpages > 9) and (mxpages > strPageNumberSize)) or ((counter > 9) and (mxpages < strPageNumberSize)) then
				ref = ref & "&nbsp;"
			end if		
			ref = ref & widenum(counter) & "<a href='topic.asp?"
            ref = ref & "TOPIC_ID=" & rs("TOPIC_ID")
            ref = ref & "&FORUM_ID=" & rs("FORUM_ID")
            ref = ref & "&CAT_ID=" & rs("CAT_ID")
            ref = ref & "&Topic_Title=" & ChkString(left(rs("T_SUBJECT"), 50),"urlpath")
            ref = ref & "&Forum_Title=" & ChkString(rs("F_SUBJECT"),"urlpath")
			ref = ref & "&whichpage=" & counter
			ref = ref & "'>" & counter & "</a></font></td>"
			Response.Write ref 
			if counter mod strPageNumberSize = 0 then
				Response.Write("</tr><tr><td>&nbsp;</td>")
			end if
		next				
        Response.Write("</tr></table>")
	end if
end sub
%>
