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

if Request.QueryString("FORUM_ID") = "" and (Request.Form("Method_Type") <> "login") and (Request.Form("Method_Type") <> "logout") then
	Response.Redirect "default.asp"
end if 
%>
<!--#INCLUDE FILE="config.asp" -->
<!--#INCLUDE FILE="inc_functions.asp" -->
<%
mypage = request("whichpage")

if mypage = "" then
	mypage = 1
end if
'############################################
sortorder = request.Form("newsort")
SortImg = request.Form("sortimg")
if sortorder <> "" then
   mypage = 1
else
   sortorder = request.Form("sort")
end if
'############################################
nDays = Request.Cookies(strCookieURL & "NumDays")
OlderThan = Request.Cookies(strCookieURL & "OlderThan") 'HuwR

if Request.form("cookie") = 1 then

	if strSetCookieToForum = "1" then
		Response.Cookies(strCookieURL & "NumDays").Path = strCookieURL
		Response.Cookies(strCookieURL & "OlderThan").Path = strCookieURL
	end if
	Response.Cookies(strCookieURL & "NumDays") = Request.Form("days")
	Response.Cookies(strCookieURL & "NumDays").expires = strForumTimeAdjust + 365
	Response.Cookies(strCookieURL & "OlderThan") = Request.Form("OlderThan")
	Response.Cookies(strCookieURL & "OlderThan").expires = strForumTimeAdjust + 365
	OlderThan = Request.Form("OlderThan")
	nDays = Request.Form("Days")
	mypage = 1
end if

if nDays = "" then
	nDays = 30
end if

defDate = DateToStr(dateadd("d", -(nDays), StrToDate(strForumTimeAdjust)))

if (strAuthType = "nt") then
	set my_Conn = Server.CreateObject("ADODB.Connection")
	my_Conn.Open strConnString
	call NTauthenticate()
	if (ChkAccountReg() = "1") then
		call NTUser()
	end if
end if

'if strPrivateForums = "1" then
'	if Request("Method_Type") = "" and (mLev <> 4) then
'		chkUser4()
'	end if
'end if
%>
<!--#INCLUDE FILE="inc_top.asp" -->
<%
'###############################################
	Dim objConn, strConn, CheckInTime, SQL, User, Timedout, Date
	Dim objRS, strSQL, houron, minon, Datec
	Dim OnlineLocation
	OnlineLocation = "<a href=forum.asp?" & Request.QueryString & ">" & Request.QueryString("Forum_Title") & "</a>"
	if Request.cookies("User")("Name") = "" then
		User = Request.ServerVariables("REMOTE_ADDR") & "Guest"
	else
		User = Request.cookies("User")("Name")
	end if
	Date = DateToStr(strForumTimeAdjust)
	CheckInTime = DateDiff( "s", CDate("01/01/1970"), Now)
	strSql = "SELECT " & strTablePrefix & "ONLINE.UserID "
	strSql = strSql & " FROM " & strTablePrefix & "ONLINE "
	strSql = strSql & " WHERE " & strTablePrefix & "ONLINE.UserID = '" & user & "'"
	set rsWho = my_Conn.Execute (strSql)
	Set objConn = Server.CreateObject("ADODB.Connection")
	objConn.ConnectionString = strConnString
	objConn.Open
	SQL = "UPDATE FORUM_ONLINE SET M_BROWSE = '" & OnlineLocation & "' , DateCreated = '" & Date & "' WHERE UserID = '" & User & "'"
	objConn.Execute SQL
	TimedOut = CheckInTime - 360 'time out the user
	SQL = "DELETE FROM FORUM_ONLINE WHERE CheckedIn < '" & TimedOut & "'"
	objConn.Execute SQL
'###############################################
if strPrivateForums = "1" then
	if Request("Method_Type") = "" and (mLev <> 4) then
		chkUser4()
	end if
end if


if (mLev = 4) or (chkForumModerator(Request.QueryString("FORUM_ID"), STRdbntUserName)= "1") or (lcase(strNoCookies) = "1") then
 	AdminAllowed = 1
else   
 	AdminAllowed = 0
end if

'## Forum_SQL - Find out if the Category is Locked or Un-Locked and if it Exists
strSql = "SELECT " & strTablePrefix & "CATEGORY.CAT_STATUS " 
strSql = strSql & " FROM " & strTablePrefix & "CATEGORY "
strSql = strSql & " WHERE " & strTablePrefix & "CATEGORY.CAT_ID = " & Request.QueryString("CAT_ID")

set rsCStatus = my_Conn.Execute (StrSql)

'## Forum_SQL - Find out if the Topic is Locked or Un-Locked and if it Exists
strSql = "SELECT " & strTablePrefix & "FORUM.F_STATUS,  " & strTablePrefix & "FORUM.F_PRIVATEFORUMS  "
strSql = strSql & " FROM " & strTablePrefix & "FORUM "
strSql = strSql & " WHERE " & strTablePrefix & "FORUM.FORUM_ID = " & Request.QueryString("FORUM_ID")

set rsFStatus = my_Conn.Execute (StrSql)
strPrivateForums =rsFStatus("F_PRIVATEFORUMS")
'## Forum_SQL - Get all topics from DB
strSql ="SELECT " & strTablePrefix & "TOPICS.T_STATUS, " & strTablePrefix & "TOPICS.CAT_ID, " 
strSql = strSql & strTablePrefix & "TOPICS.FORUM_ID, " & strTablePrefix & "TOPICS.TOPIC_ID, " 
strSql = strSql & strTablePrefix & "TOPICS.T_VIEW_COUNT, " & strTablePrefix & "TOPICS.T_SUBJECT, " 
strSql = strSql & strTablePrefix & "TOPICS.T_MAIL, " & strTablePrefix & "TOPICS.T_AUTHOR, " 
strSql = strSql & strTablePrefix & "TOPICS.T_REPLIES, " & strTablePrefix & "TOPICS.T_LAST_POST, "
strSql = strSql & strTablePrefix & "TOPICS.T_LAST_POST_AUTHOR, "  
strSql = strSql & strMemberTablePrefix & "MEMBERS.M_NAME, "
strSql = strSql & strMemberTablePrefix & "MEMBERS_1.M_NAME AS LAST_POST_AUTHOR_NAME "
strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS, "
strSql = strSql & strTablePrefix & "TOPICS, " 
strSql = strSql & strMemberTablePrefix & "MEMBERS AS " & strMemberTablePrefix & "MEMBERS_1 "
strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "TOPICS.T_AUTHOR "
strSql = strSql & " AND " & strTablePrefix & "TOPICS.T_LAST_POST_AUTHOR = "& strMemberTablePrefix & "MEMBERS_1.MEMBER_ID "
strSql = strSql & " AND " & strTablePrefix & "TOPICS.FORUM_ID = " & Request.QueryString("FORUM_ID") & " "
if nDays = "-1" then
	strSql = strSql & " AND   " & strTablePrefix & "TOPICS.T_STATUS <> 0 "
end if
if nDays > "0" then
	if OlderThan = "1" then
		strSql = strSql & " AND " & strTablePrefix & "TOPICS.T_LAST_POST < '" & defDate & "'"
	else
		strSql = strSql & " AND " & strTablePrefix & "TOPICS.T_LAST_POST > '" & defDate & "'"
	end if
end if
strSql = strSql & " ORDER BY " & strTablePrefix & "TOPICS.T_INPLACE DESC "
'############################################
select case sortorder
	case "1a", "1d"
		strSql = strSql & ",  " & strTablePrefix & "TOPICS.T_SUBJECT"
	case "2a", "2d"
		strSql = strSql & ",  " & strTablePrefix & "TOPICS.T_AUTHOR"
	case "3a", "3d"
		strSql = strSql & ",  " & strTablePrefix & "TOPICS.T_REPLIES"
	case "4a", "4d"
		strSql = strSql & ",  " & strTablePrefix & "TOPICS.T_VIEW_COUNT"
	case "5a", "5d"
		strSql = strSql & ",  " & strTablePrefix & "TOPICS.T_LAST_POST"
	case else
		strSql = strSql & ",  " & strTablePrefix & "TOPICS.T_LAST_POST"
End Select
If InStr(1, sortorder, "a", 1) Then
   strSql = strSql & " ASC"
   SortImg = "" & strImageURL & "icon_sortUP.gif"
Else
   strSql = strSql & " DESC"
   SortImg = "" & strImageURL & "icon_sortDOWN.gif"
End if
'############################################
if strDBType = "mysql" then 'MySql specific code
	if mypage > 1 then 
		intOffset = CInt((mypage-1) * strPageSize)
		strSql = strSql & " LIMIT " & intOffset & ", " & strPageSize & " "
	end if

	'## Forum_SQL - Get the total pagecount 
	strSql2 = "SELECT COUNT(" & strTablePrefix & "TOPICS.TOPIC_ID) AS PAGECOUNT "
	strSql2 = strSql2 & " FROM " & strTablePrefix & "TOPICS " 
	strSql2 = strSql2 & " WHERE   " & strTablePrefix & "TOPICS.TOPIC_ID > 0 " 
	strSql2 = strSql2 & " AND " & strTablePrefix & "TOPICS.FORUM_ID = " & Request.QueryString("FORUM_ID") & " "
	if nDays = "-1" then
		strSql2 = strSql2 & " AND   " & strTablePrefix & "TOPICS.T_STATUS <> 0 "
	end if
	if nDays > "0" then
		if OlderThan = "1" then
			strSql2 = strSql2 & " AND " & strTablePrefix & "TOPICS.T_LAST_POST < '" & defDate & "'"
		else
			strSql2 = strSql2 & " AND " & strTablePrefix & "TOPICS.T_LAST_POST > '" & defDate & "'"
		end if
	end if

	set rsCount = my_Conn.Execute(strSql2)
	if not rsCount.eof then
		maxpages = (rsCount("PAGECOUNT") \ strPageSize )
			if rsCount("PAGECOUNT") mod strPageSize <> 0 then
				maxpages = maxpages + 1
			end if
	else
		maxpages = 0
	end if 

	rsCount.close
	
	set rs = Server.CreateObject("ADODB.Recordset")
'	rs.cachesize=20

	rs.open  strSql, my_Conn, 3
	if not (rs.EOF or rs.BOF) then
		rs.movefirst
	end if
 
else 'end MySql specific code

	set rs = Server.CreateObject("ADODB.Recordset")
	rs.cachesize=20

	rs.open  strSql, my_Conn, 3
	if not (rs.EOF or rs.BOF) then
		rs.movefirst
		rs.pagesize = strPageSize
		maxpages = cint(rs.pagecount)
		rs.absolutepage = mypage
	end if
	
end if

'## Forum_SQL - Get all Forum Categories From DB
strSql = "SELECT CAT_ID FROM " & strTablePrefix & "CATEGORY"

set rsCat = my_Conn.Execute (StrSql)

%>
<script language="JavaScript">
<!----- 
function jumpTo(s) {if (s.selectedIndex != 0) top.location.href = s.options[s.selectedIndex].value;return 1;}
function setDays(order) 
{	
	if (order.search('a') > 0)
		document.DaysFilter.sortimg.value = '<%=strImageURL %>icon_sortUp.gif'
	else
		document.DaysFilter.sortimg.value = '<%=strImageURL %>icon_sortDown.gif';
	if (order != '') 
		document.DaysFilter.newsort.value = order; 
		document.DaysFilter.submit(); 
		return 0;
}
// -->
</script>
<%if rsCStatus.EOF = true OR rsFStatus.EOF = true then	Response.Redirect("default.asp")end if%> 
<table border="0" width="100%">
  <tr>
    <td width="33%" align="left" nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    <a href="default.asp"><img src="<%=strImageURL %>icon_folder_open.gif" alt="返回论坛首页" height=15 width=15 border="0"></a>&nbsp;<a href="default.asp">返回论坛首页</a>&nbsp;<img src="<% if rsCStatus("CAT_STATUS") <> 0 and rsFStatus("F_STATUS") <> 0 then Response.Write("" & strImageURL & "icon_folder_open_topic.gif") else Response.Write("" & strImageURL & "icon_folder_closed_topic.gif") %>" height=15 width=15 border="0">&nbsp;<% =ChkString(Request.QueryString("FORUM_Title"),"display") %></font>
    </font></td>
    <td align="center" width="33%"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
<% if (mlev = 4 or mlev = 3 or mlev = 2 or mlev = 1) or (lcase(strNoCookies) = "1") or (Request.Cookies(strCookieURL & "User")("Name") = "" or Request.Cookies(strCookieURL & "User")("Pword") = "") then %>
		<center><% call PostNewTopic() %></center>       
<% else %>
        &nbsp;
<% end if %>
    </font></td>
    <td align="center" width="33%">
    <form action="<% =Request.ServerVariables("SCRIPT_NAME") & "?" & Request.Querystring  %>" method="post" name="DaysFilter">
    <select name="Days" onchange="javascript:setDays('5a');">
      <option value="0" <% if ndays = "0" then Response.Write(" SELECTED")%>>显示所有主题</option>
      <option value="-1" <% if ndays = "-1" then Response.Write(" SELECTED")%>>显示所有开放的主题</option>
      <option value="1" <% if ndays = "1" then Response.Write(" SELECTED")%>>显示最近一天的主题</option>
      <option value="2" <% if ndays = "2" then Response.Write(" SELECTED")%>>显示最近两天的主题</option>
      <option value="5" <% if ndays = "5" then Response.Write(" SELECTED")%>>显示最近五天的主题</option>
      <option value="7" <% if ndays = "7" then Response.Write(" SELECTED")%>>显示最近一周的主题</option>
      <option value="14" <% if ndays = "14" then Response.Write(" SELECTED")%>>显示最近两周的主题</option>
      <option value="30" <% if ndays = "30" then Response.Write(" SELECTED")%>>显示最近一个月的主题</option>
      <option value="60" <% if ndays = "60" then Response.Write(" SELECTED")%>>显示最近两个月的主题</option>
      <option value="120" <% if ndays = "120" then Response.Write(" SELECTED")%>>显示最近四个月的主题</option>
      <option value="365" <% if ndays = "365" then Response.Write(" SELECTED")%>>显示最近一整年的主题</option>
    </select>
<% '############################################ %>
	<input type="hidden" name="newsort" value="<%=SortOrder %>">	
	<input type="hidden" name="sortimg" value="<%=SortImg %>">	
<% '############################################ %>
<% '######################## %>	
	<br><input type="Checkbox" name="OlderThan" value ="1" <% if OlderThan = "1" then Response.Write(" checked")%> onclick="javascript:setDays('5a');">显示日期之外的主题
<% '######################## %>	
    <input type="hidden" name="Cookie" value="1">
   </form>
    </td>
  </tr>
  <tr>
	<td colspan=2>
	</td>
     <td align="right">
<% if maxpages > 1 then %>
    <table border=0 align="right">
      <tr>
        <td valign="top"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>页数：</b> &nbsp;</font></td>
        <td valign="top"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% Call Paging() %></font></td>
      </tr>
    </table>
<% else %>
    &nbsp;
<% end if %>
    </td>
  </tr>
</table>
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center">
<%'############### READ/WRITE ACCESS ###########################%>
<p><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>"><b>
<%
	select case cInt(strPrivateForums)
		case 10
			response.write "只有管理员和版主可以在本论坛回复主题"
		case 11
			response.write "只有管理员和版主可以在本论坛发表新主题"
		case 12
			response.write "这是一个加密论坛，无法发表"
		case else
	end select
%>
</b></font></p>
<%'############### READ/WRITE ACCESS ###########################%>
  <tr>
    <td bgcolor="<% =strTableBorderColor %>">
    <table border="0" width="100%" cellspacing="1" cellpadding="4">
      <tr>
<% '############################################ %>
        <td align="center" bgcolor="<% =strHeadCellColor %>">
	  <b>
	     <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
		<img name="SortImg" src="<%= sortimg %>"  border="0" alt="">
	     </font>
	  </b>
	</td>
        <td align="center" bgcolor="<% =strHeadCellColor %>">
	  <b>
	     <% if sortorder = "1d" then %>
	        <a href="javascript:setDays('1a')">
	     <% else %>
	        <a href="javascript:setDays('1d')">
	     <% end if %>
	     <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">
		主 题</font></a>
	  </b><% '############################################ %>
<% if (AdminAllowed > 0) or (lcase(strNoCookies) = "1") then %>
<b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><%	call ForumAdminOptions() %></font></b><% end if %>
	</td>
        <td align="center" bgcolor="<% =strHeadCellColor %>">
	  <b>
	     <% if sortorder = "2d" then %>
	        <a href="javascript:setDays('2a')">
	     <% else %>
	        <a href="javascript:setDays('2d')">
	     <% end if %>
	     <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">
		作者</font></A>
	  </b>
	</td>
        <td align="center" bgcolor="<% =strHeadCellColor %>">
	  <b>
	     <% if sortorder = "3d" then %>
	        <a href="javascript:setDays('3a')">
	     <% else %>
	        <a href="javascript:setDays('3d')">
	     <% end if %>
	     <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">
		回复</font></A>
	  </b>
	</td>
        <td align="center" bgcolor="<% =strHeadCellColor %>">
	  <b>
	     <% if sortorder = "4d" then %>
	        <a href="javascript:setDays('4a')">
	     <% else %>
	        <a href="javascript:setDays('4d')">
	     <% end if %>
	     <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">
		人气</font></A>
	  </b>
	</td>
        <td align="center" bgcolor="<% =strHeadCellColor %>">
	  <b>
	     <% if sortorder = "5d" then %>
	        <a href="javascript:setDays('5a')">
	     <% else %>
	        <a href="javascript:setDays('5d')">
	     <% end if %>
	     <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">
		最新文章</font></A>
	  </b>
	</td>
      </tr>
<% if rs.EOF or rs.BOF then %>
      <tr>
        <td colspan="7" bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>没有找到任何主题</b></font></td>
      </tr>
<% else
	rec = 1
	do until rs.EOF or (rec = strPageSize + 1) %>
      <tr>
        <td bgcolor="<% =strForumCellColor %>" align=center valign="center"><a href="topic.asp?TOPIC_ID=<% =rs("TOPIC_ID") %>&FORUM_ID=<% =Request.QueryString("FORUM_ID") %>&CAT_ID=<% =Request.QueryString("CAT_ID") %>&Topic_Title=<% =ChkString(left(rs("T_SUBJECT"), 50),"urlpath") %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>"><%
		if rsCStatus("CAT_STATUS") <> 0 and rsFStatus("F_STATUS") <> 0 and rs("T_STATUS") <> 0 then %><% =ChkIsNew(rs("T_LAST_POST")) %><%
		else
			if rs("T_LAST_POST") > Session(strCookieURL & "last_here_date") then
				Response.Write "<img src=""" & strImageURL & "icon_folder_new_locked.gif"" height=15 width=15 border=0 hspace=0 alt=""主题已锁定"">"
			else
				Response.Write "<img src=""" & strImageURL & "icon_folder_locked.gif"" height=15 width=15 border=0 hspace=0 alt=""主题已锁定"">"
			end if
		end if %></a></td>
        <td bgcolor="<% =strForumCellColor %>" valign="center" align="left" ><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="topic.asp?TOPIC_ID=<% =rs("TOPIC_ID") %>&FORUM_ID=<% =Request.QueryString("FORUM_ID") %>&CAT_ID=<% =Request.QueryString("CAT_ID") %>&Topic_Title=<% =ChkString(left(rs("T_SUBJECT"), 50),"urlpath") %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>"><% =ChkString(left(rs("T_SUBJECT"), 30),"display") %></a>...</font><% if strShowPaging = "1" then TopicPaging() end if%><br><% 		if AdminAllowed = 1 or strNoCookies = "1" then %>
<img src="space.gif" height=0 border="0" width=255>
<%			if rsCStatus("CAT_STATUS") = 0 then %>
          <a href="JavaScript:openWindow('pop_open.asp?mode=Category&CAT_ID=<% =Request.QueryString("CAT_ID") %>')">解锁</a>
<%			else
				if rsFStatus("F_STATUS") = 0 then %>
          <a href="JavaScript:openWindow('pop_open.asp?mode=Forum&FORUM_ID=<% =Request.QueryString("FORUM_ID") %>&CAT_ID=<% =Request.QueryString("CAT_ID") %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"JSurlpath") %>')">解锁</a>
<%				else 
					if rs("T_STATUS") <> 0 then %>
          <a href="JavaScript:openWindow('pop_lock.asp?mode=Topic&TOPIC_ID=<% =rs("TOPIC_ID")%>&FORUM_ID=<% =rs("FORUM_ID") %>&CAT_ID=<% =Request.QueryString("CAT_ID") %>&Topic_Title=<% =ChkString(rs("T_SUBJECT"),"JSurlpath") %>')">锁定</a>
<%					else %>
          <a href="JavaScript:openWindow('pop_open.asp?mode=Topic&TOPIC_ID=<% =rs("TOPIC_ID")%>&FORUM_ID=<% =rs("FORUM_ID") %>&CAT_ID=<% =Request.QueryString("CAT_ID") %>&Topic_Title=<% =ChkString(rs("T_SUBJECT"),"JSurlpath") %>')">解锁</a>
<%					end if 
				end if
			end if 
			if (AdminAllowed = 1) or (rsCStatus("CAT_STATUS") <> 0 and rsFStatus("F_STATUS") <> 0 and rs("T_STATUS") <> 0) then %>
          <a href="post.asp?method=EditTopic&TOPIC_ID=<% =rs("TOPIC_ID") %>&FORUM_ID=<% =rs("FORUM_ID") %>&CAT_ID=<% =rs("CAT_ID") %>&auth=<% =ChkString(rs("T_AUTHOR"),"urlpath") %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>&Topic_Title=<% =ChkString(rs("T_SUBJECT"),"urlpath") %>">编辑</a>
<%			end if %>
          <a href="JavaScript:openWindow('pop_delete.asp?mode=Topic&TOPIC_ID=<% =rs("TOPIC_ID") %>&FORUM_ID=<% =rs("FORUM_ID") %>&CAT_ID=<% =Request.QueryString("CAT_ID") %>&Topic_Title=<% =ChkString(rs("T_SUBJECT"),"JSurlpath") %>')">删除</a>
          <a href="post.asp?method=Reply&TOPIC_ID=<% =rs("TOPIC_ID") %>&FORUM_ID=<% =rs("FORUM_ID") %>&CAT_ID=<% =Request.QueryString("CAT_ID") %>&Topic_Title=<% =ChkString(rs("T_SUBJECT"),"urlpath") %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>">回复</a>
        </td>
<%		end if %>
        <td bgcolor="<% =strForumCellColor %>" valign="center" align="center" width="10%"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strForumFontColor %>"><% =ChkString(rs("M_NAME"),"display") %></font></td>
        <td bgcolor="<% =strForumCellColor %>" valign="center" align="center" width="5%"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strForumFontColor %>"><% =rs("T_REPLIES") %></font></td>
        <td bgcolor="<% =strForumCellColor %>" valign="center" align="center" width="5%"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strForumFontColor %>"><% =rs("T_VIEW_COUNT") %></font></td>
        <%
        if IsNull(rs("T_LAST_POST_AUTHOR")) then
            strLastAuthor = ""
        else
            strLastAuthor = "<br>作者：" 
            if strUseExtendedProfile then
				strLastAuthor = strLastAuthor & "<a href=""pop_profile.asp?mode=display&id="& rs("T_LAST_POST_AUTHOR") & """>"
			else
				strLastAuthor = strLastAuthor & "<a href=""JavaScript:openWindow2('pop_profile.asp?mode=display&id=" & rs("T_LAST_POST_AUTHOR") & "')"">"
			end if
            strLastAuthor = strLastAuthor & rs("LAST_POST_AUTHOR_NAME") & "</a>"
        end if
        %>
        <td bgcolor="<% =strForumCellColor %>" valign="center" align="center" nowrap><font face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>" size="<% =strFooterFontSize %>"><b><% =ChkDate(rs("T_LAST_POST")) %></b>&nbsp<% =ChkTime(rs("T_LAST_POST")) %><%=strLastAuthor%></font></td>


      </tr>
<%		rec = rec + 1 
		rs.MoveNext 
	loop 
 end if %>
    </table>
    </td>
  </tr>
  <tr>
  <td colspan=5>
  <% if maxpages > 1 then %>
    <table border=0 align="left">
      <tr>
        <td valign="top"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">共有 <b><% =maxpages %></b> 页： </font></td>
        <td valign="top"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% Call Paging() %></font></td>
      </tr>
    </table>
<% else %>
    &nbsp;
<% end if %>
	</td>
	</tr>
</table>

<table width="100%">
  <tr>
    <td align="center" valign="top" width="33%">
    <table>
      <tr>
        <td>
		<p><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>"> 
		<img alt="新文章" src="<%=strImageURL %>icon_folder_new.gif" width="8" height="9"> 有新文章<br>
		<img alt="旧文章" src="<%=strImageURL %>icon_folder.gif" width="8" height="9"> 开放的主题<% if lcase(strHotTopic) = "1" then %>(<img alt="热门主题" src="<%=strImageURL %>icon_folder_hot.gif" width="8" height="9">超过 <% =intHotTopicNum %> 个回复的主题)<% end if %><br>
		<img alt="锁定的主题" src="<%=strImageURL %>icon_folder_locked.gif" width="8" height="9"> 锁定的主题<br>
		</font></p>
	    </td>
	  </tr>
	</table>
    </td>
    <td align="center" valign="top" width="33%">
<% if (mlev = 4 or mlev = 3 or mlev = 2 or mlev = 1) or (lcase(strNoCookies) = "1") then %>
		<p align="center"><center><% call PostNewTopic() %></center></p>       
<% else %>
        &nbsp;
<% end if %>
    </td>
    <td align="center" valign="top" width="33%">
<!--#INCLUDE FILE="inc_jump_to.asp" -->
    </td>
  </tr>
</table>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
Function ChkIsNew(dt)
	if lcase(strHotTopic) = "1" then
		if dt > Session(strCookieURL & "last_here_date") then
			if rs("T_REPLIES") >= intHotTopicNum Then
			        ChkIsNew =  "<img src='" & strImageURL & "icon_folder_new_hot.gif' height=15 width=15 border=0 hspace=0 alt='Hot Topic'>"
			else
			        ChkIsNew =  "<img src='" & strImageURL & "icon_folder_new.gif' height=15 width=15 border=0 hspace=0 alt='New Topic'>"
			end if
		Else
			if rs("T_REPLIES") >= intHotTopicNum Then
			        ChkIsNew =  "<img src='" & strImageURL & "icon_folder_hot.gif' height=15 width=15 border=0 hspace=0 alt='Hot Topic'>"
			else
			        ChkIsNew = "<img src='" & strImageURL & "icon_folder.gif' height=15 width=15 border=0 hspace=0>" 
			end if
		end if
	else
		if dt > Session(strCookieURL & "last_here_date") then
			ChkIsNew =  "<img src='" & strImageURL & "icon_folder_new.gif' height=15 width=15 border=0 hspace=0 alt='New Topic'>" 
		Else
			ChkIsNew = "<img src='" & strImageURL & "icon_folder.gif' height=15 width=15 border=0 hspace=0>" 
		end if
	end if
End Function

sub PostNewTopic() 
%>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">

<%'############### READ/WRITE ACCESS ###########################
if cInt(strPrivateForums) < 11 or AdminAllowed = 1 then
	if rsCStatus("CAT_STATUS") = 0 or rsFStatus("F_STATUS") = 0 then 
		if (AdminAllowed = 1) then %>
			<a href="post.asp?method=Topic&FORUM_ID=<% =Request.QueryString("FORUM_ID")%>&CAT_ID=<% =Request.QueryString("CAT_ID")%>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>"><img src="<%=strImageURL %>icon_folder_locked.gif" alt="锁定分类" height=15 width=15 border=0></a>&nbsp;<a href="post.asp?method=Topic&FORUM_ID=<% =Request.QueryString("FORUM_ID")%>&CAT_ID=<% =Request.QueryString("CAT_ID")%>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>">发表新主题</a>
<%		else %>
			<img src="<%=strImageURL %>icon_folder_locked.gif" alt="锁定分类" height=15 width=15 border=0>&nbsp;分类已锁定
<%		end if 
	else 
		if rsFStatus("F_STATUS") <> 0 then %>
			<a href="post.asp?method=Topic&FORUM_ID=<% =Request.QueryString("FORUM_ID")%>&CAT_ID=<% =Request.QueryString("CAT_ID")%>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>"><img src="<%=strImageURL %>icon_folder_new_topic.gif" alt="发表新主题" height=15 width=15 border=0></a>&nbsp;<a href="post.asp?method=Topic&FORUM_ID=<% =Request.QueryString("FORUM_ID")%>&CAT_ID=<% =Request.QueryString("CAT_ID")%>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>">发表新主题</a>
	<%	else %>
			<img src="<%=strImageURL %>icon_folder_locked.gif" alt="锁定论坛" height=15 width=15 border=0>&nbsp;论坛已锁定
	<%	end if 
	end if 
end if%>
</font>
<%
end sub

sub ForumAdminOptions() 
%>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
<%	if (AdminAllowed = 1) or (lcase(strNoCookies) = "1") then 
		if rsCStatus("CAT_STATUS") = 0 then 
			if mlev = 4 then %>
    <a href="JavaScript:openWindow('pop_open.asp?mode=Category&CAT_ID=<% =Request.QueryString("CAT_ID") %>')"><img src="<%=strImageURL %>icon_folder_unlocked.gif" alt="解开分类锁定" height=15 width=15 border=0></a>
<%			else %>
    <img src="<%=strImageURL %>icon_folder_locked.gif" alt="锁定分类" height=15 width=15 border=0>
<%			end if 
		else 
			if rsFStatus("F_STATUS") <> 0 then %>
    <a href="JavaScript:openWindow('pop_lock.asp?mode=Forum&FORUM_ID=<% =Request.QueryString("FORUM_ID") %>&CAT_ID=<% =Request.QueryString("CAT_ID") %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"JSurlpath") %>')"><img src="<%=strImageURL %>icon_folder_locked.gif" alt="锁定讨论区" height=15 width=15 border=0></a>
<%			else %>
    <a href="JavaScript:openWindow('pop_open.asp?mode=Forum&FORUM_ID=<% =Request.QueryString("FORUM_ID") %>&CAT_ID=<% =Request.QueryString("CAT_ID") %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"JSurlpath") %>')"><img src="<%=strImageURL %>icon_folder_unlocked.gif" alt="解开讨论区锁定" height=15 width=15 border=0></a>
<%			end if 
		end if 
		if (rsCStatus("CAT_STATUS") <> 0 and rsFStatus("F_STATUS") <> 0) or (AdminAllowed = 1) then %>
          <a href="post.asp?method=EditForum&FORUM_ID=<% =Request.QueryString("FORUM_ID") %>&CAT_ID=<% =Request.QueryString("CAT_ID") %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"JSurlpath") %>&type=0"><img src="<%=strImageURL %>icon_folder_pencil.gif" alt="编辑讨论区属性" border="0" hspace="0"></a>
<%		end if %>
    <a href="JavaScript:openWindow('pop_delete.asp?mode=Forum&FORUM_ID=<% =Request.QueryString("FORUM_ID") %>&CAT_ID=<% =Request.QueryString("CAT_ID") %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"JSurlpath") %>')"><img src="<%=strImageURL %>icon_folder_delete.gif" alt="删除讨论区" height=15 width=15 border=0></a>
    <a href="post.asp?method=Topic&FORUM_ID=<% =Request.QueryString("FORUM_ID")%>&CAT_ID=<% =Request.QueryString("CAT_ID")%>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"JSurlpath") %>"><img src="<%=strImageURL %>icon_folder_new_topic.gif" alt="发表新主题" height=15 width=15 border=0></a>
<%	end if %>
</font>
<% 
end sub

sub Paging()
	if maxpages > 1 then
		if mypage = "" then
			pge = 1
		else
			pge = mypage
		end if
		scriptname = request.servervariables("script_name")
		Response.Write("<table border=0 cellspacing=0 cellpadding=1 valign=top><tr>")
		for counter = 1 to maxpages
			if counter <> cint(pge) then   
				ref = "<td align=right bgcolor=" & strPageBGColor & "><font face=" & strDefaultFontFace & " size=" & strDefaultFontSize & ">" & "&nbsp;" & widenum(counter) & "<a href='" & scriptname 
				ref = ref & "?FORUM_ID=" & Request.QueryString("FORUM_ID") 
				ref = ref & "&CAT_ID=" & Request.QueryString("CAT_ID")
				ref = ref & "&Forum_Title=" & ChkString(Request.QueryString("FORUM_TITLE"),"urlpath") 
				ref = ref & "&whichpage=" & counter
'############################################
				if sortorder <> "" then
					ref = ref & "&sort=" & sortorder
				end if
'############################################
				ref = ref & "'>" & counter & "</a></font></td>"
				Response.Write ref 
			else
				Response.Write("<td align=right bgcolor=" & strPageBGColor & "><font face=" & strDefaultFontFace & " size=" & strDefaultFontSize & ">" & "&nbsp;" & widenum(counter) & "<b>" & counter & "</b></font></td>")
			end if
			if counter mod strPageNumberSize = 0 then
				Response.Write("</tr><tr>")
			end if
		next
		Response.Write("</tr></table>")
	end if
end sub

sub TopicPaging()
    mxpages = (rs("T_REPLIES") / strPageSize)
    if mxPages <> cint(mxPages) then
        mxpages = int(mxpages) + 1
    end if
    if mxpages > 1 then
		Response.Write("<table border=0 cellspacing=0 cellpadding=0><tr><td valign=""center""><font face=" & strDefaultFontFace & " size=" & strFooterFontSize & "><img src=""" & strImageURL & "icon_posticon.gif"" border=""0""></font></td>")
		for counter = 1 to mxpages
			ref = "<td align=right bgcolor=" & strForumCellColor  & "><font face=" & strDefaultFontFace & " size=" & strFooterFontSize & ">" 
			if ((mxpages > 9) and (mxpages > strPageNumberSize)) or ((counter > 9) and (mxpages < strPageNumberSize)) then
				ref = ref & "&nbsp;"
			end if		
			ref = ref & widenum(counter) & "<a href='topic.asp?"
            ref = ref & "TOPIC_ID=" & rs("TOPIC_ID")
            ref = ref & "&FORUM_ID=" & rs("FORUM_ID")
            ref = ref & "&CAT_ID=" & rs("CAT_ID")
            ref = ref & "&Topic_Title=" & ChkString(left(rs("T_SUBJECT"), 50),"urlpath")
            ref = ref & "&Forum_Title=" & ChkString(Request.QueryString("FORUM_Title"),"urlpath")
			ref = ref & "&whichpage=" & counter
'############################################
			if sortorder <> "" then
				ref = ref & "&sort=" & sortorder
			end if
'############################################
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
