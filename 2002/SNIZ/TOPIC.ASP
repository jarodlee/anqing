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

if Request.QueryString("TOPIC_ID") = "" and Request.QueryString("mode") <> "getIP" and Request.Form("Method_Type") <> "login" and Request.Form("Method_Type") <> "logout" then
	Response.Redirect "default.asp"
	Response.End
end if 
 %>
<!--#INCLUDE FILE="config.asp" -->
<!--#INCLUDE FILE="inc_functions.asp" -->
<%
select case Request.QueryString("mode")
	case "getIP"
 %>
<!--#INCLUDE FILE="inc_top_short.asp"-->
<%
		DisplayIP
		IP = ""
		Title = ""
 %>
<!--#INCLUDE FILE="inc_footer_short.asp" -->
<%
		Response.End
	case else
		'## Do Nothing - Continue
end select
Member_ID = getNewMemberNumber()
if (strAuthType = "nt") then
	set my_Conn = Server.CreateObject("ADODB.Connection")
	my_Conn.Open strConnString
	call NTauthenticate()
	if (ChkAccountReg() = "1") then
		call NTUser()
	end if
end if

%>
<!--#INCLUDE FILE="inc_top.asp" -->
<% 
if strPrivateForums = "1" then
	if Request("Method_Type") = "" then
		chkUser4()
	end if
end if

if (mLev = 4) or (chkForumModerator(Request.QueryString("FORUM_ID"), ChkString(STRdbntUserName, "decode"))= "1") or (lcase(strNoCookies) = "1") then
 	AdminAllowed = 1
else   
 	AdminAllowed = 0
end if


'## Forum_SQL - Find out if the Category is Locked or Un-Locked and if it Exists
strSql = "SELECT " & strTablePrefix & "CATEGORY.CAT_STATUS " 
strSql = strSql & " FROM " & strTablePrefix & "CATEGORY "
strSql = strSql & " WHERE " & strTablePrefix & "CATEGORY.CAT_ID = " & Request.QueryString("CAT_ID")

set rsCStatus = my_Conn.Execute (StrSql)

'## Forum_SQL - Find out if the Forum is Locked or Un-Locked and if it Exists
strSql = "SELECT " & strTablePrefix & "FORUM.F_STATUS, " & strTablePrefix & "FORUM.F_PRIVATEFORUMS "
strSql = strSql & " FROM " & strTablePrefix & "FORUM "
strSql = strSql & " WHERE " & strTablePrefix & "FORUM.FORUM_ID = " & Request.QueryString("FORUM_ID")

set rsFStatus = my_Conn.Execute (StrSql)
strPrivateForums =rsFStatus("F_PRIVATEFORUMS")

'## Forum_SQL - Find out if the Topic is Locked or Un-Locked and if it Exists
strSql = "SELECT " & strTablePrefix & "TOPICS.T_STATUS " 
strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
strSql = strSql & " WHERE " & strTablePrefix & "TOPICS.TOPIC_ID = " & Request.QueryString("TOPIC_ID")

set rsTStatus = my_Conn.Execute (StrSql)

if rsCStatus.EOF or rsCStatus.BOF or rsFStatus.EOF or rsFStatus.BOF or rsTStatus.EOF or rsTStatus.BOF then
	Response.Redirect(strForumURL)
else
	mypage = request("whichpage")

	if mypage = "" then
	   mypage = 1
	end if

'####################################
	sortorder = request("sortchoice")
	if sortorder <> "" then
		mypage = 1
		Application.Lock
		Application(strCookieURL & "strSortchoice") = sortorder
		Application.UnLock
	else
		sortorder = Application(strCookieURL & "strSortchoice")
	end if
'####################################

	mySQL = request("strSql")
	if mySQL = "" then
	   mySQL = SQLtemp
	end if

	'## Forum_SQL - Get all Forum Categories From DB
	strSql = "SELECT CAT_ID "
	strSql = strSql & " FROM " & strTablePrefix & "CATEGORY"

	set rsCat = my_Conn.Execute (StrSql)
	
'####################################
 %>
<script language="JavaScript">
<!--
function jumpTo(s) {if (s.selectedIndex != 0) top.location.href = s.options[s.selectedIndex].value;return 1;}
function setSort() {document.SortFilter.submit(); return 0;}
// -->
</script>
<%
'####################################

	'## Forum_SQL
	strSql ="SELECT " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.M_ICQ, " & strMemberTablePrefix & "MEMBERS.M_YAHOO, " & strMemberTablePrefix & "MEMBERS.M_AIM, " & strMemberTablePrefix & "MEMBERS.M_TITLE, " & strMemberTablePrefix & "MEMBERS.M_TITLE, " & strMemberTablePrefix & "MEMBERS.M_Homepage, " & strMemberTablePrefix & "MEMBERS.M_LEVEL, " & strMemberTablePrefix & "MEMBERS.M_POSTS, " & strMemberTablePrefix & "MEMBERS.M_COUNTRY, " & strTablePrefix & "REPLY.REPLY_ID, " & strTablePrefix & "REPLY.R_AUTHOR, " & strTablePrefix & "REPLY.TOPIC_ID, " & strTablePrefix & "REPLY.R_MESSAGE, " & strTablePrefix & "REPLY.R_DATE "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS, " & strTablePrefix & "REPLY "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "REPLY.R_AUTHOR "
	strSql = strSql & " AND   TOPIC_ID = " & Request.QueryString("TOPIC_ID") & " "
	strSql = strSql & " ORDER BY " & strTablePrefix & "REPLY.R_DATE"

	if strDBType = "mysql" then 'MySql specific code

		'## Forum_SQL - Get the total pagecount 
		strSql2 = "SELECT COUNT(" & strTablePrefix & "REPLY.TOPIC_ID) AS REPLYCOUNT "
		strSql2 = strSql2 & " FROM " & strMemberTablePrefix & "MEMBERS, " & strTablePrefix & "REPLY "
		strSql2 = strSql2 & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "REPLY.R_AUTHOR "
		strSql2 = strSql2 & " AND   TOPIC_ID = " & Request.QueryString("TOPIC_ID") & " "
		
		set rsCount = my_Conn.Execute(strSql2)
		if not rsCount.eof then
			maxpages = (rsCount("REPLYCOUNT")  \ strPageSize )
			if rsCount("REPLYCOUNT") mod strPageSize <> 0 then
				maxpages = maxpages + 1
			end if
		else
			maxpages = 1
		end if 
	
		set rs = Server.CreateObject("ADODB.Recordset")
'		rs.cachesize= strPageSize

		rs.open  strSql, my_Conn, 3
		if not(rs.EOF) then
			rs.movefirst
		end if
		
	else 'end MySql specific code
	
		set rs = Server.CreateObject("ADODB.Recordset")

		rs.cachesize = strPageSize
		rs.open  strSql, my_Conn, 3

		If not (rs.EOF or rs.BOF) then  '## No replies found in DB
			rs.movefirst
			rs.pagesize = strPageSize
			rs.absolutepage = mypage '**
			maxpages = cint(rs.pagecount)
		end if
	end if
	i = 0 
 %>
<table border="0" width="100%">
  <tr>
	<td width="100%" align="left" nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
	<img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="default.asp">返回论坛首页</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="FORUM.asp?FORUM_ID=<% =Request.QueryString("FORUM_ID") %>&CAT_ID=<% =Request.QueryString("CAT_ID") %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>"><% =ChkString(Request.QueryString("FORUM_Title"),"display") %></a>
<% call PostingOptions() %>
    </font></td><form action="<% =Request.ServerVariables("SCRIPT_NAME") & "?" & Request.Querystring  %>" method="post" name="SortFilter">
<td align=right> 
<select name="SortChoice" onchange="javascript:setSort();">
      <option value="0" <% if sortorder = "0" then Response.Write(" SELECTED")%>>旧文章排前面</option>
      <option value="1" <% if sortorder = "1" then Response.Write(" SELECTED")%>>新文章排前面</option>
      <option value="2" <% if sortorder = "2" then Response.Write(" SELECTED")%>>先显示原文章再从最新排起</option>
    </select>
   
  </td></form></tr>
<% '#################################### %>
  <tr>
  <table border="0" width="100%" cellspacing="0" cellpadding="3" align="center">
  <tr>
     <td align=left>

<%'############### READ/WRITE ACCESS ###########################%>

<br><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>
<%
	select case cInt(strPrivateForums)
		case 10
			response.write "只有管理者和版主可以在本讨论区回覆主题"
		case 11
			response.write "只有管理者和版主可以在本讨论区发表新主题"
		case 12
			response.write "这是唯读讨论区，无法发表"
		case else
	end select
%>
</b></font>
<%'############### READ/WRITE ACCESS ###########################%>
	 </td>
<td align="left"><a href="print.asp?<%=Request.QueryString%>"><img border="0" src="<%=strImageURL %>print.gif" border=0></a> <a href="print.asp?<%=Request.QueryString%>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize-1 %>">把本主题输出到打印机</font></a>
	 <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize-1 %>">
	 <a href="bookmark.asp?mode=add&id=<%=Request.querystring("Topic_ID")%>"><img border="0" src="<%=strImageURL %>icon_bookmark.gif" border=0></a> <a href="bookmark.asp?mode=add&id=<%=Request.querystring("Topic_ID")%>">将本主题添加进收藏夹</a></font></td>
<td align="right">
<% if maxpages > 1 then %>
    <table border=0 align="right">
      <tr>
        <td valign="top"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>页数：</b></font></td>
        <td valign="top"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% Call Paging() %></font></td>
      </tr>
    </table>
<% else %>
	<td align=right>&nbsp;</td>
    &nbsp;
<% end if %>
    </td>
	</tr>
	</table>
  </tr>
</table>

<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td bgcolor="<% =strTableBorderColor %>">
    <table border="0" width="100%" cellspacing="1" cellpadding="4">
      <tr>
        <td align="center" bgcolor="<% =strHeadCellColor %>" width="<% =strTopicWidthLeft %>" <% if lcase(strTopicNoWrapLeft) = "1" then Response.Write(" nowrap") %>><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">作者</font></b></td>
        <td align="left" bgcolor="<% =strHeadCellColor %>" width="<% =strTopicWidthRight %>" <% if lcase(strTopicNoWrapRight) = "1" then Response.Write(" nowrap") %>><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>"><% if strShowTopicNav = "1" then Call Topic_nav()  else Response.Write("主题") end if%></font></b></td>

      </tr>
<% 
'####################################
	if (sortorder <> "1") and MyPage = 1 then 
		Call GetFirst() 
	end if
'####################################

	'## Forum_SQL - Get all topicsFrom DB
	strSql ="SELECT " & strMemberTablePrefix & "MEMBERS.M_NAME, " 
	strsql = strsql & strMemberTablePrefix & "MEMBERS.M_PHOTO_URL, "
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_ICQ, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_YAHOO, "
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_AIM, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_TITLE, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_Homepage, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_LEVEL, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_POSTS, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_PMRECEIVE, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_COUNTRY, " 
	strSql = strSql & strTablePrefix & "REPLY.REPLY_ID, " 
	strSql = strSql & strTablePrefix & "REPLY.R_AUTHOR, " 
	strSql = strSql & strTablePrefix & "REPLY.TOPIC_ID, " 
	strSql = strSql & strTablePrefix & "REPLY.R_MESSAGE, " 
	strSql = strSql & strTablePrefix & "REPLY.R_DATE "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS, " & strTablePrefix & "REPLY "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "REPLY.R_AUTHOR "
	strSql = strSql & " AND   TOPIC_ID = " & Request.QueryString("TOPIC_ID") & " "
'####################################
	if sortorder = "2" then
		strSql = strSql & " ORDER BY R_DATE DESC"
	else
		if sortorder = "1" then
			strSql = strSql & " ORDER BY R_DATE DESC"
		else
			strSql = strSql & " ORDER BY R_DATE ASC"
		end if
	end if
'####################################

	
	if strDBType = "mysql" then 'MySql specific code
		if mypage > 1 then 
			intOffSet = CInt((mypage - 1) * strPageSize) - 1
			strSql = strSql & " LIMIT " & intOffSet & ", " & CInt(strPageSize) & " "
		end if

		'## Forum_SQL - Get the total pagecount 
		strSql2 = "SELECT COUNT(" & strTablePrefix & "REPLY.TOPIC_ID) AS REPLYCOUNT "
		strSql2 = strSql2 & " FROM " & strTablePrefix & "REPLY "
		strSql2 = strSql2 & " WHERE  TOPIC_ID = " & Request.QueryString("TOPIC_ID") & " "
		
		set rsCount = my_Conn.Execute(strSql2)
		if not rsCount.eof then
			maxpages = (rsCount("REPLYCOUNT")  \ strPageSize )
			if rsCount("REPLYCOUNT") mod strPageSize <> 0 then
				maxpages = maxpages + 1
			end if
		else
			maxpages = 1
		end if 

		set rs = Server.CreateObject("ADODB.Recordset")
'		rs.cachesize = strPageSize

		rs.open  strSql, my_Conn, 3

	else 'end MySql specific code
	
		set rs = Server.CreateObject("ADODB.Recordset")
		rs.cachesize = 20
		rs.open  strSql, my_Conn, 3
	
		if not(rs.EOF or rs.BOF) then  '## Replies found in DB
			rs.movefirst
			rs.pagesize = strPageSize
			maxpages = cint(rs.pagecount)
			rs.absolutepage = mypage
		end if
	end if		
	if rs.EOF or rs.BOF then  '## No replies found in DB
		Response.Write ""
	else
		'rs.movefirst			
		intI = 0 
		howmanyrecs = 0
		rec = 1
	
		do until rs.EOF or (mypage = 1 and rec > strPageSize) or (mypage > 1 and rec > strPageSize) '**		
			if intI = 0 then 
				CColor = strAltForumCellColor
			else
				CColor = strForumCellColor
			end if
 %>
      <tr>
        <td bgcolor="<% =CColor %>" valign="top" align="center">
        <% if strUseExtendedProfile then %>
		<a href="pop_profile.asp?mode=display&id=<% =rs("R_AUTHOR") %>">
        <% else %>
		<a href="JavaScript:openWindow3('pop_profile.asp?mode=display&id=<% =rs("R_AUTHOR") %>')">
		<% end if %>	
		<font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b><% =ChkString(rs("M_NAME"),"display") %></a>
        </b></font>
        
        <% if Trim(rs("M_PHOTO_URL")) <> "" and lcase(rs("M_PHOTO_URL")) <> "http://" then %>
        <br><img src="<% =rs("M_PHOTO_URL") %>" align="center">
        <% end if %>

<%				if strShowRank = 1 or strShowRank = 3 then %>
        <br><font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>"><% = ChkString(getMember_Level(rs("M_TITLE"), rs("M_LEVEL"), rs("M_POSTS")),"display") %></font>
<%				end if %>
<%				if strShowRank = 2 or strShowRank = 3 then %>
        <br><% = getStar_Level(rs("M_LEVEL"), rs("M_POSTS")) %>
<%				end if %>
        <br>
        <br><font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>"><% =rs("M_COUNTRY") %></font>
        <br><font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">发表总数：<% =rs("M_POSTS") %></font></td>
        <td bgcolor="<% =CColor %>" <% if (AdminAllowed = 1) then %>colspan="3"<% else %>colspan="2"<% end if %> valign="top"><img src="<%=strImageURL %>icon_posticon.gif" border="0" hspace="3"><font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">发表於&nbsp;-&nbsp;<% =ChkDate(rs("R_DATE")) %>&nbsp;:&nbsp;<% =ChkTime(rs("R_DATE")) %></font>
        <% if strUseExtendedProfile then %>
		&nbsp;<a href="pop_profile.asp?mode=display&id=<% =rs("MEMBER_ID") %>"><img src="<%=strImageURL %>1-profile.gif" alt="查看会员资料" border="0" align="absmiddle"></a>
        <% else %>
		&nbsp;<a href="JavaScript:openWindow3('pop_profile.asp?mode=display&id=<% =rs("MEMBER_ID") %>')"><img src="<%=strImageURL %>1-profile.gif" alt="查看会员资料" border="0" align="absmiddle"></a>
		<% end if %>	
   
<%		if (lcase(strEmail) = "1") then 
			if (mlev <> 0) or (mlev = 0 and  strLogonForMail <> "1") then 
%> <a href="JavaScript:openWindow('pop_mail.asp?id=<% =rs("MEMBER_ID") %>')"><img src="<%=strImageURL %>1-email.gif" alt="电子邮件" border="0" align="absmiddle"></a>
<%			end if
		else
%> <a href="JavaScript:openWindow('pop_mail.asp?id=<% =rs("MEMBER_ID") %>')"><img src="<%=strImageURL %>1-email.gif" alt="电子邮件" border="0" align="absmiddle"></a>
<%		end if %> 
<%		if (intPMessages = 1) and (rs("M_PMRECEIVE") = "1") then %>
 <a href="privatesend.asp?method=Topic&mname=<% =rs("M_NAME") %>"><img src="<%=strImageURL %>1-pm.gif" alt="发送悄悄话" border="0" align="absmiddle"></a>
<%		end if  %>

<%			if strHomepage = "1" then %>
<%				if rs("M_Homepage") <> " " then %>
 <a href="<% =rs("M_Homepage") %>"><img src="<%=strImageURL %>1-home.gif" alt="浏览 <% = ChkString(rs("M_NAME"),"display") %> 的主页" border="0" align="absmiddle"></a>
<%				end if %>
<%			end if %>
<%			if (AdminAllowed = 1 or rs("MEMBER_ID") = Member_ID) or (strNoCookies = "1") then %>
<%				if (rsCStatus("CAT_STATUS") <> 0 and rsFStatus("F_STATUS") <> 0 and rsTStatus("T_STATUS") <> 0) or (AdminAllowed = 1) then %>
 <a href="post.asp?method=Edit&REPLY_ID=<% =rs("REPLY_ID") %>&TOPIC_ID=<% =Request.QueryString("TOPIC_ID") %>&FORUM_ID=<% =Request.QueryString("FORUM_ID") %>&CAT_ID=<% =Request.QueryString("CAT_ID") %>&auth=<% =ChkString(rs("R_AUTHOR"),"urlpath") %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>&Topic_Title=<% =ChkString(Request.QueryString("Topic_Title"),"urlpath") %>"><img src="<%=strImageURL %>1-edit.gif" alt="编辑讯息" border="0" align="absmiddle"></a>
<%				end if %>
<%			end if %>
<%			if strICQ = "1" then %>
<%			  if Trim(rs("M_ICQ")) <> "" then %>
 <a href="JavaScript:openWindow('pop_messengers.asp?mode=ICQ&ICQ=<% =ChkString(rs("M_ICQ"), "JSurlpath") %>&M_NAME=<% =ChkString(rs("M_NAME"),"JSurlpath") %>')"><img src="http://online.mirabilis.com/scripts/online.dll?icq=<% =ChkString(rs("M_ICQ"), "urlpath") %>&img=5" height=15 width=15 alt="发送ICQ讯息给 <% = ChkString(rs("M_NAME"),"display")  %>" border="0" align="absmiddle"></a>
<%			  end if %>
<%			end if %>
<%			if strYAHOO = "1" then %>
<%			  if Trim(rs("M_YAHOO")) <> "" then %>
 <a href="http://search.tencent.com/cgi-bin/friend/user_show_info?ln=<% =rs("M_YAHOO") %>"  target=_blank><img align=absmiddle width=16 height=16 src="http://icon.tencent.com/<% =rs("M_YAHOO") %>/s/00/" alt="<% =rs("M_YAHOO") %>" border=0></a> <% =rs("M_YAHOO") %>
<%			  end if %>
<%			end if %>
<%			if (strAIM = "1") then %>
<%				if Trim(rs("M_AIM")) <> "" then %>
        &nbsp;<a href="JavaScript:openWindow('pop_messengers.asp?mode=AIM&AIM=<% =rs("M_AIM") %>&M_NAME=<% =ChkString(rs("M_NAME"),"urlpath") %>')"><img src="<%=strImageURL %>icon_aim.gif" height=15 width=15 alt="传AIM讯息给 <% =rs("M_NAME") %>" border="0" align="absmiddle" hspace="6"></a>
<%				end if %>
<%			end if %>
<%'############### READ/WRITE ACCESS ###########################%>
<%			if (rsCStatus("CAT_STATUS") <> 0 and rsFStatus("F_STATUS") <> 0 and rsTStatus("T_STATUS") <> 0) or (AdminAllowed = 1) then
			if (cint(strPrivateForums) <> 10 and cint(strPrivateForums) <> 12) or AdminAllowed = 1 then %>
 <a href="post.asp?method=ReplyQuote&REPLY_ID=<% =rs("REPLY_ID") %>&TOPIC_ID=<% =Request.QueryString("TOPIC_ID") %>&FORUM_ID=<% =Request.QueryString("FORUM_ID") %>&CAT_ID=<% =Request.QueryString("CAT_ID") %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>&Topic_Title=<% =ChkString(Request.QueryString("Topic_Title"),"urlpath") %>&M=<% =Request.QueryString("M") %>"><img src="<%=strImageURL %>1-quote.gif" alt="回复并引用原文" border="0" align="absmiddle"></a>
<%			end if 
			end if%>
<%'############### READ/WRITE ACCESS ###########################%>			
<%			if (strIPLogging = "1") then %>
<%				if (AdminAllowed = 1) or (strNoCookies = "1") then %>
 <a href="JavaScript:openWindow('topic.asp?mode=getIP&REPLY_ID=<% =rs("REPLY_ID") %>&FORUM_ID=<% =Request.QueryString("FORUM_ID") %>')"><img src="<%=strImageURL %>icon_ip.gif" height=15 width=15 alt="查看会员IP" border="0" align="absmiddle"></a>
<%				end if %>
<%			end if %>
<%			if (AdminAllowed = 1 or rs("MEMBER_ID") = Member_ID) or (strNoCookies = "1") then %>
<%				if (rsCStatus("CAT_STATUS") <> 0 and rsFStatus("F_STATUS") <> 0 and rsTStatus("T_STATUS") <> 0) or (AdminAllowed = 1) then %>
 <a href="JavaScript:openWindow('pop_delete.asp?mode=Reply&REPLY_ID=<% =Rs("REPLY_ID") %>&TOPIC_ID=<% =Request.QueryString("TOPIC_ID") %>&FORUM_ID=<% =Request.QueryString("FORUM_ID") %>')"><img src="<%=strImageURL %>icon_delete_reply.gif" height=15 width=15 alt="删除回复" border="0" align="absmiddle"></a>
<%				end if %>
<%			end if %>

        <hr noshade size="<% =strFooterFontSize %>">
        
        <font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% =formatStr(rs("R_MESSAGE")) %><a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" height=15 width=15 border="0" align="right" alt="回到顶部"></a></font></td>
      </tr>
<%			
		    rs.MoveNext
		    intI  = intI + 1
		    if intI = 2 then 
				intI = 0
			end if
		    rec = rec + 1
		loop
	end if
'####################################
		if sortorder = "1" then 
		Call GetFirst() 
		end if
'####################################
	
 %>
    </table></td>
  </tr>
  <tr>
    <td colspan="2">
    <table border="0" width="100%">
      <tr>
        <td>
<% if maxpages > 1 then %>
        <table border=0>
          <tr>
            <td valign="top"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>主题共有 <% =maxpages %> 页：</font></td>
            <td valign="top"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% Call Paging() %></font></td>
          </tr>
        </table>
<% else %>
	<td valign="top">&nbsp;</td>
<% end if %>
        </td>
        <td align="right" nowrap>
<% if (AdminAllowed = 1) then %>
<%	call AdminOptions() %>
<% else %>
        &nbsp;
<% end if %>
        </td> 
      </tr>
    </table></td>
  </tr>
</table>
</div>

<table width="100%">
  <tr>
    <td align="left" valign="top" width="50%"><%	 Call PostingOptions() %></td>
    <td align="right" valign="top" width="50%"><!--#INCLUDE FILE="inc_jump_to.asp" --></td>
  </tr>
</table>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
end if

sub GetFirst()

	'## Forum_SQL - Get Origional Posting
	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.M_NAME, " 	& strMemberTablePrefix & "MEMBERS.M_ICQ, "  & strMemberTablePrefix & "MEMBERS.M_PMRECEIVE, "  & strMemberTablePrefix & "MEMBERS.M_PHOTO_URL, "
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_YAHOO, " & strMemberTablePrefix & "MEMBERS.M_AIM, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_TITLE, " & strMemberTablePrefix & "MEMBERS.M_HOMEPAGE, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_LEVEL, " 
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_POSTS, " & strMemberTablePrefix & "MEMBERS.M_COUNTRY, " 
	strSql = strSql & strTablePrefix & "TOPICS.T_DATE, " & strTablePrefix & "TOPICS.T_SUBJECT, " & strTablePrefix & "TOPICS.T_AUTHOR, " 
	strSql = strSql & strTablePrefix & "TOPICS.TOPIC_ID, " & strTablePrefix & "TOPICS.T_MESSAGE "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS, " & strTablePrefix & "TOPICS "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "TOPICS.T_AUTHOR "
	strSql = strSql & " AND   " & strTablePrefix & "TOPICS.TOPIC_ID = " &  Request.QueryString("TOPIC_ID") 

	set rs = my_Conn.Execute (strSql)

	if rs.EOF or rs.BOF then  '## No categories found in DB
		Response.Write "  <tr>" & vbCrLf
		Response.Write "    <td colspan=5>没有找到任何主题</td>" & vbCrLf
		Response.Write "  </tr>" & vbCrLf
	else
 %>
      <tr>
        <td bgcolor="<% =strForumFirstCellColor %>" valign="top" align="center">
        <font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
        <% if strUseExtendedProfile then %>
		<a href="pop_profile.asp?mode=display&id=<% =rs("MEMBER_ID") %>">
        <% else %>
		<a href="JavaScript:openWindow3('pop_profile.asp?mode=display&id=<% =rs("MEMBER_ID") %>')">
		<% end if %>	
		<font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b><% =ChkString(rs("M_NAME"),"display") %></a>
        </b></font>

        <% if Trim(rs("M_PHOTO_URL")) <> "" and lcase(rs("M_PHOTO_URL")) <> "http://" then %>
        <br><img src="<% =rs("M_PHOTO_URL") %>">
        <% end if %>

<%		if strShowRank = 1 or strShowRank = 3 then %>
        <br><font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>"><% = ChkString(getMember_Level(rs("M_TITLE"), rs("M_LEVEL"), rs("M_POSTS")),"display") %></font>
<%		end if %>
<%		if strShowRank = 2 or strShowRank = 3 then %>
        <br><% = getStar_Level(rs("M_LEVEL"), rs("M_POSTS")) %>
<%		end if %>
        <br>
        <br><font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>"><% =rs("M_COUNTRY") %></font>
        <br><font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">发表总数：<% =rs("M_POSTS") %></font></td>
        <td bgcolor="<% =strForumFirstCellColor %>" <% if (AdminAllowed = 1) then %>colspan="3"<% else %>colspan="2"<% end if %> valign="top"><img src="<%=strImageURL %>icon_posticon.gif" border="0" hspace="3"><font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">发表於&nbsp;-&nbsp;<% =ChkDate(rs("T_DATE")) %> <% =ChkTime(rs("T_DATE")) %></font>

        <% if strUseExtendedProfile then %>
		&nbsp;<a href="pop_profile.asp?mode=display&id=<% =rs("MEMBER_ID") %>"><img src="<%=strImageURL %>1-profile.gif" alt="查看会员资料" border="0" align="absmiddle"></a>
        <% else %>
		&nbsp;<a href="JavaScript:openWindow3('pop_profile.asp?mode=display&id=<% =rs("MEMBER_ID") %>')"><img src="<%=strImageURL %>1-profile.gif" alt="查看示会员资料" border="0" align="absmiddle"></a>
		<% end if %>	
        
<%		if (lcase(strEmail) = "1") then 
			if (mlev <> 0) or (mlev = 0 and  strLogonForMail <> "1") then 
%> <a href="JavaScript:openWindow('pop_mail.asp?id=<% =rs("MEMBER_ID") %>')"><img src="<%=strImageURL %>1-email.gif" alt="电子邮件" border="0" align="absmiddle"></a>
<%			end if
		else
%> <a href="JavaScript:openWindow('pop_mail.asp?id=<% =rs("MEMBER_ID") %>')"><img src="<%=strImageURL %>1-email.gif" alt="电子邮件" border="0" align="absmiddle"></a>
<%		end if %>        
<%		if (intPMessages = 1) and (rs("M_PMRECEIVE") = "1") then %>
 <a href="privatesend.asp?method=Topic&mname=<% =rs("M_NAME") %>"><img src="<%=strImageURL %>1-pm.gif" alt="发送悄悄话" border="0" align="absmiddle"></a>
<%		end if  %>
<%		if (strHomepage = "1") then %>
<%				if rs("M_Homepage") <> " " then %>
 <a href="<% =rs("M_Homepage") %>"><img src="<%=strImageURL %>1-home.gif"  alt="参观 <% =ChkString(rs("M_NAME"),"display")  %> '的个人主页" border="0" align="absmiddle"></a>
<%			end if %>
<%		end if %>
<%		if (AdminAllowed = 1 or rs("MEMBER_ID") = Member_ID) or (strNoCookies = "1") then %>
<%			if ((rsCStatus("CAT_STATUS") <> 0) and (rsFStatus("F_STATUS") <> 0) and (rsTStatus("T_STATUS") <> 0)) or (AdminAllowed = 1) then %>
 <a href="post.asp?method=EditTopic&REPLY_ID=<% =rs("TOPIC_ID") %>&TOPIC_ID=<% =Request.QueryString("TOPIC_ID") %>&FORUM_ID=<% =Request.QueryString("FORUM_ID") %>&CAT_ID=<% =Request.QueryString("CAT_ID") %>&auth=<% =rs("T_AUTHOR") %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>&Topic_Title=<% =ChkString(Request.QueryString("Topic_Title"),"urlpath") %>"><img src="<%=strImageURL %>1-edit.gif" alt="编辑讯息" border="0" align="absmiddle"></a>
<%			end if %>
<%		end if %>
<%		if (strICQ = "1") then %>
<%			if Trim(rs("M_ICQ")) <> "" then %>
 <a href="JavaScript:openWindow('pop_messengers.asp?mode=ICQ&ICQ=<% =ChkString(rs("M_ICQ"), "JSurlpath") %>&M_NAME=<% =ChkString(rs("M_NAME"), "JSurlpath") %>')"><img src="http://online.mirabilis.com/scripts/online.dll?icq=<% =ChkString(rs("M_ICQ"), "urlpath") %>&img=5" height=16 width=16 alt="发送ICQ讯息给 <% =ChkString(rs("M_NAME"),"display")  %>" border="0" align="absmiddle"></a>
<%			end if %>
<%		end if %>
<%		if (strYAHOO = "1") then %>
<%		  if Trim(rs("M_YAHOO")) <> "" then %>
 <a href="http://search.tencent.com/cgi-bin/friend/user_show_info?ln=<% =rs("M_YAHOO") %>"  target=_blank><img align=absmiddle width=16 height=16 src="http://icon.tencent.com/<% =rs("M_YAHOO") %>/s/00/" alt="<% =rs("M_YAHOO") %>" border=0></a> <% =rs("M_YAHOO") %>
<%		  end if %>
<%		end if %>
<%		if (strAIM = "1") then %>
<%			if Trim(rs("M_AIM")) <> "" then %>
        &nbsp;<a href="JavaScript:openWindow('pop_messengers.asp?mode=AIM&AIM=<% =rs("M_AIM") %>&M_NAME=<% =rs("M_NAME") %>')"><img src="<%=strImageURL %>icon_aim.gif" height=16 width=16 alt="传AIM讯息给 <% =rs("M_NAME") %>" border="0" align="absmiddle" hspace="6"></a>
<%			end if %>
<%		end if %>
<%'############### READ/WRITE ACCESS ###########################%>
<%			if (rsCStatus("CAT_STATUS") <> 0 and rsFStatus("F_STATUS") <> 0 and rsTStatus("T_STATUS") <> 0) or (AdminAllowed = 1) then
			if (cint(strPrivateForums) <> 10 and cint(strPrivateForums) <> 12) or AdminAllowed = 1 then %>
 <a href="post.asp?method=TopicQuote&TOPIC_ID=<% =Request.QueryString("TOPIC_ID") %>&FORUM_ID=<% =Request.QueryString("FORUM_ID") %>&CAT_ID=<% =Request.QueryString("CAT_ID") %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>&Topic_Title=<% =ChkString(Request.QueryString("Topic_Title"),"urlpath") %>"><img src="<%=strImageURL %>1-quote.gif" alt="引用回复" border="0" align="absmiddle"></a>
<%			end if 
			end if%>
<%'############### READ/WRITE ACCESS ###########################%>
<%		if (strIPLogging = "1") then %>
<%			if (AdminAllowed = 1) or (strNoCookies = "1") then %>
 <a href="JavaScript:openWindow('topic.asp?mode=getIP&TOPIC_ID=<% =rs("TOPIC_ID") %>&FORUM_ID=<% =Request.QueryString("FORUM_ID") %>')"><img src="<%=strImageURL %>icon_ip.gif" height=16 width=16 alt="查看会员IP" border="0" align="absmiddle" hspace="6"></a>
<%			end if %>
<%		end if %>
        <hr noshade size="1">
        
        <font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% =formatStr(rs("T_MESSAGE")) %></font></td>
      </tr>
<%
	end if

	'## Forum_SQL
	strSql = "UPDATE " & strTablePrefix & "TOPICS "
	strSql = strSql & " SET " & strTablePrefix & "TOPICS.T_VIEW_COUNT = (" & strTablePrefix & "TOPICS.T_VIEW_COUNT + 1) "
	strSql = strSql & " WHERE (" & strTablePrefix & "TOPICS.TOPIC_ID = " & Request.QueryString("TOPIC_ID") & ");"

	my_conn.Execute (strSql)

	set rs = nothing

End Sub
sub DisplayIP()
	usr = (chkForumModerator(Request.QueryString("FORUM_ID"), STRdbntUserName))
	if (chkUser((STRdbntUserName), (Request.Cookies(strUniqueID & "User")("Pword"))) = 4) then 
		usr = 1
	end if
	if usr then
		if Request.QueryString("TOPIC_ID") <> "" then

			'## Forum_SQL
			strSql = "SELECT " & strTablePrefix & "TOPICS.T_IP, " & strTablePrefix & "TOPICS.T_SUBJECT "
			strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
			strSql = strSql & " WHERE TOPIC_ID = " & Request.QueryString("TOPIC_ID")

			rsIP = my_Conn.Execute(strSql)

			IP = rsIP("T_IP")
			Title = rsIP("T_Subject")
		else
			if Request.QueryString("REPLY_ID") <> "" then
				'## Forum_SQL
				strSql = "SELECT " & strTablePrefix & "REPLY.R_IP "
				strSql = strSql & " FROM " & strTablePrefix & "REPLY "
				strSql = strSql & " WHERE REPLY_ID = " & Request.QueryString("REPLY_ID")

				rsIP = my_Conn.Execute(strSql)

				IP = rsIP("R_IP")
			end if
		end if
		set rsIP = nothing
 %>
<P align=center><b>查看会员IP地址：</b><br>
<% =ip %></P>
<%	else %>
<p align=center><b>只有版主和管理员才能执行此功能</B></p>
<%
	end If
end sub
sub PostingOptions() 
 %>
    <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
<%

	if (mlev = 4 or mlev = 3 or mlev = 2 or mlev = 1) or (lcase(strNoCookies) = "1") or (STRdbntUserName = "") then
'############### READ/WRITE ACCESS ###########################
	if cInt(strPrivateForums) < 11 or AdminAllowed = 1 then
		if ((rsCStatus("CAT_STATUS") = 1) and (rsFStatus("F_STATUS") = 1)) then %>
    <a href="post.asp?method=Topic&FORUM_ID=<% =Request.QueryString("FORUM_ID") %>&CAT_ID=<% =Request.QueryString("CAT_ID") %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>"><img src="<%=strImageURL %>icon_folder_new_topic.gif" height=15 width=15 border=0></a>&nbsp;<a href="post.asp?method=Topic&FORUM_ID=<% =Request.QueryString("FORUM_ID") %>&CAT_ID=<% =Request.QueryString("CAT_ID") %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>">发表新主题</a>
<%		else
			if (AdminAllowed = 1) then %>
    <a href="post.asp?method=Topic&FORUM_ID=<% =Request.QueryString("FORUM_ID") %>&CAT_ID=<% =Request.QueryString("CAT_ID") %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>"><img src="<%=strImageURL %>icon_folder_locked.gif" height=15 width=15 border=0></a>&nbsp;<a href="post.asp?method=Topic&FORUM_ID=<% =Request.QueryString("FORUM_ID") %>&CAT_ID=<% =Request.QueryString("CAT_ID") %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>">发表新主题</a>
<%			else %>
    <img src="<%=strImageURL %>icon_folder_locked.gif" height=15 width=15 border=0>&nbsp;论坛已锁定
<%			end if 
		end if 
	end if
'############### READ/WRITE ACCESS ###########################
	if (cint(strPrivateForums) <> 10 and cint(strPrivateForums) <> 12) or AdminAllowed = 1 then
		if (rsCStatus("CAT_STATUS") = 1) and (rsFStatus("F_STATUS") = 1) and (rsTStatus("T_STATUS") = 1) then %>
    <a href="post.asp?method=Reply&TOPIC_ID=<% =Request.QueryString("TOPIC_ID") %>&FORUM_ID=<% =Request.QueryString("FORUM_ID") %>&CAT_ID=<% =Request.QueryString("CAT_ID") %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>&Topic_Title=<% =ChkString(Request.QueryString("Topic_Title"),"urlpath") %>"><img src="<%=strImageURL %>icon_reply_topic.gif" height=15 width=15 border=0></a>&nbsp;<a href="post.asp?method=Reply&TOPIC_ID=<% =Request.QueryString("TOPIC_ID") %>&FORUM_ID=<% =Request.QueryString("FORUM_ID") %>&CAT_ID=<% =Request.QueryString("CAT_ID") %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>&Topic_Title=<% =ChkString(Request.QueryString("Topic_Title"),"urlpath") %>">回复主题</a>
<%		Else 
			if (AdminAllowed = 1)  then %>
    <a href="post.asp?method=Reply&TOPIC_ID=<% =Request.QueryString("TOPIC_ID") %>&FORUM_ID=<% =Request.QueryString("FORUM_ID") %>&CAT_ID=<% =Request.QueryString("CAT_ID") %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>&Topic_Title=<% =ChkString(Request.QueryString("Topic_Title"),"urlpath") %>"><img src="<%=strImageURL %>icon_closed_topic.gif" height=15 width=15 border=0></a>&nbsp;<a href="post.asp?method=Reply&TOPIC_ID=<% =Request.QueryString("TOPIC_ID") %>&FORUM_ID=<% =Request.QueryString("FORUM_ID") %>&CAT_ID=<% =Request.QueryString("CAT_ID") %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>&Topic_Title=<% =ChkString(Request.QueryString("Topic_Title"),"urlpath") %>">回复主题</a>
<%			Else %>
    <img src="<%=strImageURL %>icon_closed_topic.gif" height=15 width=15 border=0>&nbsp;主题已锁定
<%			end if
		end if 
	end if 
		if (lcase(strEmail) = "1") then 
			if (mlev <> 0) or (mlev = 0 and  strLogonForMail <> "1") then %>
				
				<a href="JavaScript:openWindow('pop_send_to_friend.asp?url=<% = strForumURL & "link.asp&TOPIC_ID=" & Request.QueryString("TOPIC_ID")  %>')"><img border="0" src="<%=strImageURL %>icon_send_topic.gif" height=15 width=15 border=0></a>&nbsp;<a href="JavaScript:openWindow('pop_send_to_friend.asp?url=<% = strForumURL & "link.asp&TOPIC_ID=" & Request.QueryString("TOPIC_ID")  %>')">将本主题寄给朋友</a>
				<!-- <a href="JavaScript:openWindow('pop_send_to_friend.asp?url=http://<% =Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("URL") & "?" & Request.QueryString %>')"><img border="0" src="<%=strImageURL %>icon_send_topic.gif" height=15 width=15 border=0></a>&nbsp;<a href="JavaScript:openWindow('pop_send_to_friend.asp?http://<% =Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("URL") & "?" & Request.QueryString %>')">将本主题寄给朋友</a> -->
<%			end if
		end if %>
<%	end if %>
    </font>
<% 
end sub 
sub AdminOptions() 
 %>
    <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
<%	if (AdminAllowed = 1) or (lcase(strNoCookies) = "1") then
		if (rsCStatus("CAT_STATUS") = 0) then 
			if (mlev = 4) then %>
    <a href="JavaScript:openWindow('pop_open.asp?mode=Category&CAT_ID=<% =Request.QueryString("CAT_ID") %>')"><img border="0" src="<%=strImageURL %>icon_folder_unlocked.gif" alt="解开分类锁定" height=15 width=15 border=0></a>
<%			else %>
    <img border="0" src="<%=strImageURL %>icon_folder_unlocked.gif" alt="分类已锁定" height=15 width=15 border=0>
<%			end if
		else 
			if (rsFStatus("F_STATUS") = 0) then %>
    <a href="JavaScript:openWindow('pop_open.asp?mode=Forum&FORUM_ID=<% =Request.QueryString("FORUM_ID") %>&CAT_ID=<% =Request.QueryString("CAT_ID") %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"JSurlpath") %>')"><img border="0" src="<%=strImageURL %>icon_folder_unlocked.gif" alt="解开论坛锁定" height=15 width=15 border=0></a>
<%			else
				if (rsTStatus("T_STATUS") <> 0) then %>
    <a href="JavaScript:openWindow('pop_lock.asp?mode=Topic&TOPIC_ID=<% =Request.QueryString("TOPIC_ID") %>&FORUM_ID=<% =Request.QueryString("FORUM_ID") %>&CAT_ID=<% =Request.QueryString("CAT_ID") %>&Topic_Title=<% =ChkString(Request.QueryString("Topic_Title"),"JSurlpath") %>')"><img border="0" src="<%=strImageURL %>icon_folder_locked.gif" alt="锁定论坛" height=15 width=15 border=0></a>
<%				else %>
    <a href="JavaScript:openWindow('pop_open.asp?mode=Topic&TOPIC_ID=<% =Request.QueryString("TOPIC_ID") %>&FORUM_ID=<% =Request.QueryString("FORUM_ID") %>&CAT_ID=<% =Request.QueryString("CAT_ID") %>&Topic_Title=<% =ChkString(Request.QueryString("Topic_Title"),"JSurlpath") %>')"><img border="0" src="<%=strImageURL %>icon_folder_unlocked.gif" alt="解开主题锁定" height=15 width=15 border=0></a>
<%				end if
			end if
		end if %>
<%		if ((rsCStatus("CAT_STATUS") <> 0) and (rsFStatus("F_STATUS") <> 0) and (rsTStatus("T_STATUS") <> 0)) or (AdminAllowed = 1) then %>
    <a href="post.asp?method=EditTopic&REPLY_ID=<% =Request.QueryString("TOPIC_ID") %>&TOPIC_ID=<% =Request.QueryString("TOPIC_ID") %>&FORUM_ID=<% =Request.QueryString("FORUM_ID") %>&CAT_ID=<% =Request.QueryString("CAT_ID") %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>&Topic_Title=<% =ChkString(Request.QueryString("Topic_Title"),"urlpath") %>"><img src="<%=strImageURL %>icon_folder_pencil.gif" alt="编辑主题" border="0" hspace="0"></a>
<%		end if %>
    <a href="JavaScript:openWindow('pop_delete.asp?mode=Topic&TOPIC_ID=<% =Request.QueryString("TOPIC_ID") %>&FORUM_ID=<% =Request.QueryString("FORUM_ID") %>&CAT_ID=<% =Request.QueryString("CAT_ID") %>&Topic_Title=<% =ChkString(Request.QueryString("Topic_Title"),"JSurlpath") %>')"><img border="0" src="<%=strImageURL %>icon_folder_delete.gif" alt="删除主题" height=15 width=15 border=0></a>
    <a href="post.asp?method=Topic&FORUM_ID=<% =Request.QueryString("FORUM_ID") %>&CAT_ID=<% =Request.QueryString("CAT_ID") %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>"><img src="<%=strImageURL %>icon_folder_new_topic.gif" alt="发表新主题" height=15 width=15 border=0></a>
    <a href="post.asp?method=Reply&TOPIC_ID=<% =Request.QueryString("TOPIC_ID") %>&FORUM_ID=<% =Request.QueryString("FORUM_ID") %>&CAT_ID=<% =Request.QueryString("CAT_ID") %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>&Topic_Title=<% =ChkString(Request.QueryString("Topic_Title"),"urlpath") %>"><img src="<%=strImageURL %>icon_reply_topic.gif" alt="回复主题" height=15 width=15 border=0></a>
<%	end if %>
    </font>
<% 
end sub 
sub Paging()
	if maxpages > 1 then
		if Request.QueryString("whichpage") = "" then
			pge = 1
		else
			pge = Request.QueryString("whichpage")
		end if
		scriptname = request.servervariables("script_name")
		Response.Write("<table border=0 width=100% cellspacing=0 cellpadding=1 align=top><tr>")
		for counter = 1 to maxpages
			if counter <> cint(pge) then   
				ref = "<td align=right ><font face=" & strDefaultFontFace & " size=" & strDefaultFontSize & ">" & "&nbsp;" & widenum(counter) & "<a href='" & scriptname 
				ref = ref & "?whichpage=" & counter
				'ref = ref & "&pagesize=" & mypagesize 

'####################################
				if sortorder = "1" then
					ref = ref & "&sort=" & sortorder
				else
					if sortorder = "2" then
						ref = ref & "&sort=" & sortorder
					end if
				end if
'####################################

				ref = ref & "&Forum_Title=" & ChkString(Request.QueryString("FORUM_Title"),"urlpath") 
				ref = ref & "&Topic_Title=" & ChkString(Request.QueryString("Topic_Title"),"urlpath")
				ref = ref & "&CAT_ID=" & Request.QueryString("CAT_ID")
				ref = ref & "&FORUM_ID=" & Request.QueryString("FORUM_ID") 
				ref = ref & "&TOPIC_ID=" & Request.QueryString("TOPIC_ID") & "'"
				if top = "1" then
					ref = ref & ">"
					ref = ref & "</font><b><font face='" & strDefaultFontFace & "' "
					ref = ref & "color='" & strHeadFontColor & "'"
					ref = ref & ">" & counter & "</font></b></a></td>"
					Response.Write ref
				else
					ref = ref & "'>" & counter & "</font></a></td>"
					Response.Write ref
				end if
			else
				Response.Write("<td align=right bgcolor=" & strPageBGColor & "><font face=" & strDefaultFontFace & " size=" & strDefaultFontSize & ">" & "&nbsp;" & widenum(counter) & "<b>" & counter & "</b></font></td>")
			end if
			if counter mod strPageNumberSize = 0 then
				Response.Write("</tr><tr>")
			end if
		next
		Response.Write("</tr></table>")
	end if
	top = "0"
end sub 

Sub Topic_nav()    
    set rsLastPost = Server.CreateObject("ADODB.Recordset")
    
    strSql = "SELECT T_LAST_POST FROM " & strTablePrefix & "TOPICS " 
    strSql = strSql & "WHERE TOPIC_ID = " & Request.QueryString("TOPIC_ID")
                
    set rsLastPost = my_Conn.Execute (StrSql)
    
    T_LAST_POST = rsLastPost("T_LAST_POST")
    
    strSQL = "SELECT T_SUBJECT, TOPIC_ID "
    strSql = strSql & "FROM " & strTablePrefix & "TOPICS "
    strSql = strSql & "WHERE T_LAST_POST > '" & T_LAST_POST
    strSql = strSql & "' AND FORUM_ID=" & Request.QueryString("Forum_ID")
    strSql = strSql & " ORDER BY T_LAST_POST;"
                
    set rsPrevTopic = my_conn.Execute (strSQL)
    
    strSQL = "SELECT T_SUBJECT, TOPIC_ID "
    strSql = strSql & "FROM " & strTablePrefix & "TOPICS "
    strSql = strSql & "WHERE T_LAST_POST < '" & T_LAST_POST
    strSql = strSql & "' AND FORUM_ID=" & Request.QueryString("Forum_ID")
    strSql = strSql & " ORDER BY T_LAST_POST DESC;"
                
    set rsNextTopic = my_conn.Execute (strSQL)
    
    if rsPrevTopic.EOF then
        prevTopic = "<img src=""" & strImageURL & "icon_blank.gif"" height=15 width=15 alt=""上一个主题"" border=""0"" align=""absmiddle"" hspace=""6"">"
    else
        prevTopic = "<a href=topic.asp?cat_id=" & request.queryString("CAT_ID") & _
                    "&FORUM_ID=" & Request.QueryString("FORUM_ID") & _
                    "&TOPIC_ID=" & rsPrevTopic("TOPIC_ID") & _        
                    "&Topic_Title=" & ChkString(rsPrevTopic("T_SUBJECT"),"urlpath") & _
                    "&Forum_Title=" & ChkString(Request.QueryString("Forum_Title"),"urlpath") & _
                    "><img src=""" & strImageURL & "1-prev.gif"" alt=""上一个主题"" border=""0"" align=""absmiddle"" hspace=""6""></a>"
    end if                    
                    
    if rsNextTopic.EOF then
        NextTopic = "<img src=""" & strImageURL & "icon_blank.gif"" height=15 width=15 alt=""上一个主题"" border=""0"" align=""absmiddle"" hspace=""6"">"
    else
        NextTopic = "<a href=topic.asp?cat_id=" & request.queryString("CAT_ID") & _
                    "&FORUM_ID=" & Request.QueryString("FORUM_ID") & _
                    "&TOPIC_ID=" & rsNextTopic("TOPIC_ID") & _        
                    "&Topic_Title=" & ChkString(rsNextTopic("T_SUBJECT"),"urlpath") & _
                    "&Forum_Title=" & ChkString(Request.QueryString("Forum_Title"),"urlpath") & _
                    "><img src=""" & strImageURL & "1-next.gif"" alt=""下一个主题"" border=""0"" align=""absmiddle"" hspace=""6""></a>"
    end if                    
    
    Response.Write (prevTopic & "<b><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>主题：" & Request.QueryString("Topic_Title") & "&nbsp;</font></b>" & nextTopic)
    
    rsLastPost.close
    rsPrevTopic.close
    rsNextTopic.close
    set rsLastPost = nothing
    set rsPrevTopic = nothing
    set rsNextTopic = nothing
    
end sub

 %>
