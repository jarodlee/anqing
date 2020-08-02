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
strDBNTUserName = Request.Cookies(strUniqueID & "User")("Name")
strDBNTFUserName = Request.Form("Name")
if strAuthType = "nt" then
	strDBNTUserName = Session(strCookieURL & "userID")
	strDBNTFUserName = Session(strCookieURL & "userID")
end if

if strAutoLogon = 1 then
	if (ChkAccountReg() <> "1") then
		Response.Redirect "register.asp?mode=DoIt"
	end if
end if

if IsEmpty(Session(strCookieURL & "last_here_date")) then
	Session(strCookieURL & "last_here_date") = ReadLastHereDate(strDBNTUserName)
end if


	set rs1 = Server.CreateObject("ADODB.Recordset")

	'## Forum_SQL
	strSql = "SELECT " & strTablePrefix & "TOTALS.P_COUNT, " & strTablePrefix & "TOTALS.T_COUNT, " & strTablePrefix & "TOTALS.U_COUNT "
	strSql = strSql & " FROM " & strTablePrefix & "TOTALS"

	rs1.open strSql, my_Conn

	Users = rs1("U_COUNT")
	Topics = rs1("T_COUNT")
	Posts = rs1("P_COUNT")

	rs1.Close
	set rs1 = nothing


'## Declare Variables 
Dim rsCat_ID, rsCat_Name, rsCat_Status, rsCat_Description, rsCat_SumForum, rsCat_SumTopic, rsCat_SumPost, rsCat_LastPost

'## Forum_SQL - Get all Forums From DB - mod 12/17/2000 drv
strSql = "SELECT " & strTablePrefix & "CATEGORY.CAT_ID, " &_
		" " & strTablePrefix & "CATEGORY.CAT_NAME, " &_ 
		" " & strTablePrefix & "CATEGORY.CAT_STATUS, " &_ 
		" COUNT(" & strTablePrefix & "FORUM.F_SUBJECT) AS SumForum, " &_ 
		" SUM(" & strTablePrefix & "FORUM.F_TOPICS) AS SumTopic, " &_ 
		" SUM(" & strTablePrefix & "FORUM.F_COUNT) AS SumPost, " &_ 
		" MAX(" & strTablePrefix & "FORUM.F_LAST_POST) AS LastPost " &_
	" FROM " & strTablePrefix & "FORUM RIGHT OUTER JOIN " &_
		" " & strTablePrefix & "CATEGORY ON " &_ 
		" " & strTablePrefix & "FORUM.CAT_ID = " & strTablePrefix & "CATEGORY.CAT_ID " &_
	" GROUP BY " & strTablePrefix & "CATEGORY.CAT_NAME, " & strTablePrefix & "CATEGORY.CAT_ID, " & strTablePrefix & "CATEGORY.CAT_STATUS " &_
	" ORDER BY " & strTablePrefix & "CATEGORY.CAT_NAME;"
	
'response.write strSql
'response.end

set rs = my_Conn.Execute (strSql)
%>

<table border=0 width="95%" cellspacing=0 cellpadding=0 align=center>
  <tr>
    <td>
<%
ShowLastHere = (cint(ChkUser2(strDBNTUserName, Request.Cookies(strUniqueID & "User")("Pword"))) > 0)
if strShowStatistics <> "1" then
%>
    <table width="100%" border="0">
      <tr>
        <td>
<%
	if ShowLasthere then 
%>
        <font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">上次登陆时间 - <% =ChkDate(Session(strCookieURL & "last_here_date")) %> <% =ChkTime(Session(strCookieURL & "last_here_date")) %></font>
<%
	else
%>
        &nbsp;
<%
	end if
%>
        </td>
        <td align=right><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">共 <% =Topics %> 个主题，文章 <% =Posts %> 篇，注册会员总数：  <% =Users %>&nbsp;&nbsp;</font></td>
      </tr>
    </table>
<%
else
	Response.Write("&nbsp;&nbsp;")
end if
%>
    </td>  
  </tr>
  <tr>
    <td bgcolor="<% =strTableBorderColor %>">
    <table border=0 width="100%" cellspacing=1 cellpadding=4>
      <tr>
        <td align=center bgcolor="<% =strHeadCellColor %>" nowrap valign="top"><b>&nbsp;</td>
        <td align=center bgcolor="<% =strHeadCellColor %>" nowrap valign="top"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">Forum Categories</font></b></td>
        <td align=center bgcolor="<% =strHeadCellColor %>" nowrap valign="top"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">论坛列表</font></b></td>
        <td align=center bgcolor="<% =strHeadCellColor %>" nowrap valign="top"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">主题</font></b></td>        <td align=center bgcolor="<% =strHeadCellColor %>" nowrap valign="top"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">文章</font></b></td>
        <td align=center bgcolor="<% =strHeadCellColor %>" nowrap valign="top"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">最后发表</font></b></td>
<% 
if (strShowModerators = "1") or (mlev = 4 or mlev = 3) then 
%>
        <td align=center bgcolor="<% =strHeadCellColor %>" nowrap valign="top"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">版主</font></b></td>
<%
end if 
if (mlev = 4 or mlev = 3) or (lcase(strNoCookies) = "1") then %>
        <td align=center bgcolor="<% =strHeadCellColor %>"><% PostingOptions() %></td>
<%
end if
%>
      </tr>
<% 
if rs.EOF or rs.BOF then
%>
      <tr>
        <td bgcolor="<% =strCategoryCellColor %>" colspan="<% if (strShowModerators = "1") or (mlev = 4 or mlev = 3) then Response.Write("7") else Response.Write("6")%>"><font face="<% =strDefaultFontFace %>" color="<% =strCategoryFontColor %>" size="<% =strDefaultFontSize %>" valign="top"><b>没有找到分类或论坛</b></font></td>
<%
	if (mlev = 4 or mlev = 3) then 
%>
        <td bgcolor="<% =strCategoryCellColor %>"><font face="<% =strDefaultFontFace %>" color="<% =strCategoryFontColor %>" size="<% =strDefaultFontSize %>" valign="top">&nbsp;</font></td>
<%
	end if 
%>
      </tr>
<%
else
	intPostCount  = 0
	intTopicCount = 0
	intForumCount = 0
	strLastPostDate = ""
	do until rs.EOF 


		' Set Recordset Values to Variables
		rsCat_ID = rs.Fields("CAT_ID").Value
		rsCat_Name = rs.Fields("CAT_NAME").Value
		rsCat_Status = rs.Fields("CAT_STATUS").Value
		'rsCat_Description = rs.Fields("CAT_DESCRIPTION").Value
		rsCat_SumForum = rs.Fields("SumForum").Value
		rsCat_SumTopic = rs.Fields("SumTopic").Value
		rsCat_SumPost = rs.Fields("SumPost").Value
		rsCat_LastPost = rs.Fields("LastPost").Value

		If IsNumeric(rsCat_SumPost) Then intPostCount  = intPostCount + rsCat_SumPost
		If IsNumeric(rsCat_SumTopic) Then intTopicCount = intTopicCount + rsCat_SumTopic
		If IsNumeric(rsCat_SumForum) Then intForumCount = intForumCount + rsCat_SumForum
		If (rsCat_LastPost > strLastPostDate) Then strLastPostDate = rsCat_LastPost

		Response.Write	"      <tr>" & vbcrlf

		Response.Write "        <td bgcolor=""" & strForumCellColor & """ align=""center"" valign=""top"" nowrap>" & vbCrLf
		if rsCat_Status = 0 then 
			Response.Write "        <a href=""category.asp?CAT_ID=" & rsCat_ID & """>"
			if rsCat_LastPost > Session(strCookieURL & "last_here_date") then
				Response.Write "<img src=""<%=strImageURL %>icon_folder_new_locked.gif"" height=15 width=15 border=0 hspace=0 alt=""分类已锁定""></a>"
			else
				Response.Write "<img src=""<%=strImageURL %>icon_folder_locked.gif"" height=15 width=15 border=0 hspace=0 alt=""分类已锁定""></a>"
			end if     
		else 
			Response.Write "        <a href=""category.asp?CAT_ID=" & rsCat_ID & """>" & ChkIsNew(rsCat_LastPost) & "</a>"
		end if
		Response.Write "</td>" & vbCrLf		
		Response.Write "        <td bgcolor=""" & strForumCellColor & """ align=""left"" valign=""top"" nowrap>" & vbCrLf
		Response.Write "<font face=""" & strDefaultFontFace & """ color=""" & strForumFontColor & """ size=""" & strDefaultFontSize & """><a href=""category.asp?CAT_ID=" & rsCat_ID & """>" & ChkString(rs("CAT_NAME"),"display") & "</a></font></td>" & vbcrlf 

		' Category Forum Count
		If IsNull(rsCat_SumForum) Then rsCat_SumForum = 0
		Response.Write ("<td bgcolor=""" & strForumCellColor & """ align=""center"" valign=""top"">" &_
		"<font face=""" & strDefaultFontFace & """ color=""" & strForumFontColor & """ size=""" & strDefaultFontSize & """>" &_
		rsCat_SumForum & "</font></td>" & vbCrLf)

		' Category Topic Count
		If IsNull(rsCat_SumTopic) Then rsCat_SumTopic = 0
		Response.Write ("<td bgcolor=""" & strForumCellColor & """ align=""center"" valign=""top"">" &_
		"<font face=""" & strDefaultFontFace & """ color=""" & strForumFontColor & """ size=""" & strDefaultFontSize & """>" &_
		rsCat_SumTopic & "</font></td>" & vbCrLf)

		' Category Post Count
		If IsNull(rsCat_SumPost) Then rsCat_SumPost = 0
		Response.Write ("<td bgcolor=""" & strForumCellColor & """ align=""center"" valign=""top"">" &_
		"<font face=""" & strDefaultFontFace & """ color=""" & strForumFontColor & """ size=""" & strDefaultFontSize & """>" &_
		rsCat_SumPost & "</font></td>" & vbCrLf)

		' Category Last Post Date
		Response.Write ("<td bgcolor=""" & strForumCellColor & """ align=""center"" valign=""top"" nowrap>")
		Response.Write ("<font face=""" & strDefaultFontFace & """ color=""" & strForumFontColor & """ size=""" & strFooterFontSize & """>")
		If (rsCat_LastPost <> "") Then
		Response.Write ("<b>" & ChkDate(rsCat_LastPost) & "</b>" & "<br>" & ChkTime(rsCat_LastPost))
		Else
		Response.Write ("&nbsp;")
		End If
		Response.Write ("</td><td bgcolor=""" & strForumCellColor & """ align=""center"" valign=""top"" nowrap>&nbsp;</td>")

		if (mlev = 4 or mlev = 3) or (lcase(strNoCookies) = "1") then 
			Response.Write "        <td bgcolor=""" & strForumCellColor & """ align=""center"" valign=""top"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>"
			call CategoryAdminOptions()
			Response.Write "</font></b></td>" & vbcrlf 
		end if 
		Response.Write "      </tr>" & vbcrlf
		chkDisplayHeader = false
						
		rs.MoveNext
	loop
end if 
if strShowStatistics = "1" then%>
	<!--#include file="statistics.asp"-->
<%end if %>
<!--#include file="online2.asp"-->
    </table>           
    </td>
  </tr>
  <tr>
    <td>
    <table width="100%">
      <tr>
        <td>
        <font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
        <img alt="新文章" src="<%=strImageURL %>icon_folder_new.gif" width=8 height=9> 自从上次访问后有新文章<br>
        <img alt="旧文章" src="<%=strImageURL %>icon_folder.gif" width=8 height=9> 自从上次访问后无新文章<br>
        </font>
        </td>
      </tr>
    </table>
    </td>
  </tr>
</table>

<%
set rs = nothing 
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<% 
sub PostingOptions() 
	if (mlev = 4) or (lcase(strNoCookies) = "1") then 
		Response.Write "<font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""post.asp?method=Category""><img border=0 src=""<%=strImageURL %>icon_folder_new_topic.gif"" alt=""Create New Category"" height=15 width=15 border=0></a></font>" & vbcrlf 
	else
		Response.Write "&nbsp;" & vbcrlf 
	end if 
end sub 

function ChkIsNew(dt)
	if rsCat_Status <> 0 then
		if dt > Session(strCookieURL & "last_here_date") then
			ChkIsNew =  "<IMG src='<%=strImageURL %>icon_folder_new.gif' height=15 width=15 border=0 hspace=0 alt='New Posts'>" 
		Else
			ChkIsNew = "<IMG src='<%=strImageURL %>icon_folder.gif' height=15 width=15 border=0 hspace=0 alt='Old Posts'>" 
		end if
	else
		ChkIsNew = "<IMG src='<%=strImageURL %>icon_folder_locked.gif' height=15 width=15 border=0 hspace=0 alt='论坛已锁定'>"
	end if
end function

sub CategoryAdminOptions() 
	if (mlev = 4) or (lcase(strNoCookies) = "1") then 
		if (rs("CAT_STATUS") <> 0) then 
			Response.Write "          <a href=""JavaScript:openWindow('pop_lock.asp?mode=Category&CAT_ID=" & rs("CAT_ID") & "&Cat_Title=" & ChkString(rs("CAT_NAME"),"jsURLPath") & "')""><img src=""<%=strImageURL %>icon_lock.gif"" alt=""Lock Category"" border=0 hspace=0></a>" & vbcrlf
		else
			Response.Write "          <a href=""JavaScript:openWindow('pop_open.asp?mode=Category&CAT_ID=" & rs("CAT_ID") & "')""><img src=""<%=strImageURL %>icon_unlock.gif"" alt=""Un-Lock Category"" border=0 hspace=0></a>" & vbcrlf
		end if 
		if (rs("CAT_STATUS") <> 0) or (mlev = 4) then
			Response.Write "          <a href=""post.asp?method=EditCategory&CAT_ID=" & rs("CAT_ID") & "&Cat_Title=" & ChkString(rs("CAT_NAME"),"urlpath") & """><img src=""<%=strImageURL %>icon_pencil.gif"" alt=""Edit Category Name"" border=0 hspace=0></a>" & vbcrlf
		end if
			Response.Write "          <a href=""JavaScript:openWindow('pop_delete.asp?mode=Category&CAT_ID=" & rs("CAT_ID") & "&Cat_Title=" & ChkString(rs("CAT_NAME"),"JSurlpath") & "')""><img src=""<%=strImageURL %>icon_trashcan.gif"" alt=""Delete Category"" border=0 hspace=0></a>" & vbcrlf
		if (rs("CAT_STATUS") <> 0) or (mlev = 4) then
			Response.Write "          <a href=""post.asp?method=Forum&CAT_ID=" & rs("CAT_ID") & "&type=0""><img src=""<%=strImageURL %>icon_folder_new_topic.gif"" alt=""Create New Forum"" border=0 hspace=0></a>" & vbcrlf
		end if 
		if (rs("CAT_STATUS") <> 0) or (mlev = 4) then
			Response.Write "          <a href=""post.asp?method=URL&CAT_ID=" & rs("CAT_ID") & "&type=1""><img src=""<%=strImageURL %>icon_url.gif"" alt=""Create New Web Link"" border=0 hspace=0></a>" & vbcrlf
		end if 
	else
		Response.Write "          &nbsp;" & vbcrlf
	end if
end sub 

sub ForumAdminOptions() 
	if (mLev = 4) or (chkForumModerator(rsForum("FORUM_ID"), strDBNTUserName) = "1") or (lcase(strNoCookies) = "1") then
		if rsForum("F_TYPE") = 0 then
			if rs("CAT_STATUS") = 0 then
				if (mlev = 4) then 
%>
          <a href="JavaScript:openWindow('pop_open.asp?mode=Category&CAT_ID=<% =rs("CAT_ID") %>')"><img src="<%=strImageURL %>icon_unlock.gif" alt="解开分类锁定" border="0" hspace="0"></a>
<%
				end if
			else 
				if rsForum("F_STATUS") = 1 then 
%>
          <a href="JavaScript:openWindow('pop_lock.asp?mode=Forum&FORUM_ID=<% =rsForum("FORUM_ID") %>&CAT_ID=<% =rsForum("CAT_ID") %>&Forum_Title=<% =ChkString(rsForum("F_SUBJECT"),"JSurlpath")%>')"><img src="<%=strImageURL %>icon_lock.gif" alt="锁定论坛" border="0" hspace="0"></a>
<%
				else 
%>
          <a href="JavaScript:openWindow('pop_open.asp?mode=Forum&FORUM_ID=<% =rsForum("FORUM_ID") %>&CAT_ID=<% =rsForum("CAT_ID") %>&Forum_Title=<% =ChkString(rsForum("F_SUBJECT"),"JSurlpath")%>')"><img src="<%=strImageURL %>icon_unlock.gif" alt="解开论坛锁定" border="0" hspace="0"></a>
<%
				end if 
			end if
		end if
		if rsForum("F_TYPE") = 0 then
			if (rs("CAT_STATUS") <> 0 and rsForum("F_STATUS") <> 0) or (mlev = 4 or mlev = 3) then 
%>
          <a href="post.asp?method=EditForum&FORUM_ID=<% =rsForum("FORUM_ID") %>&CAT_ID=<% =rsForum("CAT_ID") %>&Forum_Title=<% =ChkString(rsForum("F_SUBJECT"),"urlpath") %>&type=0"><img src="<%=strImageURL %>icon_pencil.gif" alt="编辑论坛属性" border="0" hspace="0"></a>
<%
			end if
		else 
			if rsForum("F_TYPE") = 1 then 
%>
          <a href="post.asp?method=EditURL&FORUM_ID=<% =rsForum("FORUM_ID") %>&CAT_ID=<% =rsForum("CAT_ID") %>&Forum_Title=<% =ChkString(rsForum("F_SUBJECT"),"urlpath") %>&type=1"><img src="<%=strImageURL %>icon_pencil.gif" alt="编辑连接属性" border="0" hspace="0"></a>
<%
			end if 
		end if 
		if (mlev = 4) or (lcase(strNoCookies) = "1") then 
%>
          <a href="JavaScript:openWindow('pop_delete.asp?mode=Forum&FORUM_ID=<% =rsForum("FORUM_ID") %>&CAT_ID=<% =rsForum("CAT_ID") %>&Forum_Title=<% =ChkString(rsForum("F_SUBJECT"),"JSurlpath") %>')"><img src="<%=strImageURL %>icon_trashcan.gif" alt="删除讨论区" border="0" hspace="0"></a>
<%
		end if
		if rsForum("F_TYPE") = 0 then
			if (mlev = 4) or (lcase(strNoCookies) = "1") then 
%>
          <a href="post.asp?method=Topic&FORUM_ID=<% =rsForum("FORUM_ID") %>&CAT_ID=<% =rsForum("CAT_ID") %>&Forum_Title=<% =ChkString(rsForum("F_SUBJECT"),"urlpath") %>"><img src="<%=strImageURL %>icon_folder_new_topic.gif" alt="发表新主题" height=15 width=15 border=0></a>
<%
			end if
		end if 
	else
		Response.Write "&nbsp;"
	end if
end sub 

sub WriteStatistics() 
	Dim Forum_Count
	Dim NewMember_Name, NewMember_Id, Member_Count
	Dim LastPostDate, LastPostLink

	set rs = Server.CreateObject("ADODB.Recordset")
	
	Forum_Count = intForumCount

	'## Forum_SQL - Get newest membername and id from DB
	strSql = "SELECT M_NAME, MEMBER_ID FROM " & strMemberTablePrefix & "MEMBERS WHERE M_STATUS=1 AND MEMBER_ID > 1"
	strSql = strSQL & " ORDER BY M_DATE desc;"
	set rs = my_Conn.Execute(strSql)
	if not rs.EOF then
		NewMember_Name = ChkString(rs("M_NAME"), "display") 
		NewMember_Id = rs("MEMBER_ID")
	else 
		NewMember_Name = ""
	end if
    
	'## Forum_SQL - Get Active membercount from DB 
	strSql = "SELECT COUNT(MEMBER_ID) AS U_COUNT FROM " & strMemberTablePrefix & "MEMBERS WHERE M_POSTS > 0 AND M_STATUS=1"
	
	set rs = my_Conn.Execute(strSql)
	
	if not rs.EOF then
		Member_Count = rs("U_COUNT")
	else
		Member_Count = 0
	end if
	
	LastPostDate = ""
 	LastPostLink = ""
	LastPostAuthorLink = ""
	
	if not (intLastPostForum_ID = "") then	
		'## Forum_SQL - Get lastPostDate and link to that post from DB
		strSql = "SELECT " & strTablePrefix & "FORUM.CAT_ID, " & strTablePrefix & "FORUM.FORUM_ID, " 
		strSql = strSql & strTablePrefix & "FORUM.F_SUBJECT, " & strTablePrefix & "TOPICS.TOPIC_ID, " & strTablePrefix & "TOPICS.T_SUBJECT, "
		strSql = strSql & strTablePrefix & "TOPICS.T_LAST_POST, " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.MEMBER_ID "
		strSql = strSql & "FROM " & strTablePrefix & "FORUM, " & strTablePrefix & "TOPICS, "
		strSql = strSql & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " WHERE " & strTablePrefix & "FORUM.FORUM_ID = " & strTablePrefix & "TOPICS.FORUM_ID "
		strSql = strSql & " AND " & strTablePrefix & "FORUM.CAT_ID = " & strTablePrefix & "TOPICS.CAT_ID "
		strSql = strSql & " AND " & strTablePrefix & "TOPICS.T_LAST_POST_AUTHOR = " & strMemberTablePrefix & "MEMBERS.MEMBER_ID "
		strSql = strSql & " AND " & strTablePrefix & "FORUM.FORUM_ID = " & intLastPostForum_ID & " "
		strSql = strSql & "ORDER BY " & strTablePrefix & "TOPICS.T_LAST_POST DESC;"

	 	set rs = my_Conn.Execute(strSql)
 	
 		if not rs.EOF then
			LastPostDate = ChkDate(rs("T_LAST_POST")) & ChkTime(rs("T_LAST_POST"))
			LastPostLink = "topic.asp?TOPIC_ID=" & rs("TOPIC_ID") & "&FORUM_ID=" & rs("FORUM_ID") & "&CAT_ID=" & rs("CAT_ID")
			LastPostLink = LastPostLink  & "&Topic_Title=" & ChkString(rs("T_SUBJECT"),"urlpath")
			LastPostLink = LastPostLink  & "&Forum_Title=" & ChkString(rs("F_SUBJECT"),"urlpath") 
			LastPostAuthorLink = " 作者："
			strMember_ID = rs("MEMBER_ID")
			strM_NAME = ChkString(rs("M_NAME"),"display") 
			if strUseExtendedProfile then
				LastPostAuthorLink = LastPostAuthorLink & "<a href=""pop_profile.asp?mode=display&id="& strMember_ID & """>"
			else
				LastPostAuthorLink = LastPostAuthorLink & "<a href=""JavaScript:openWindow2('pop_profile.asp?mode=display&id=" & strMember_ID & "')"">"
			end if
            		LastPostAuthorLink = LastPostAuthorLink  & strM_NAME & "</a>"
		end if
	end if

	ActiveTopicCount = -1
	if not IsNull(Session(strCookieURL & "last_here_date")) then 
		if not blnHiddenForums then
			'## Forum_SQL - Get ActiveTopicCount from DB
			strSql = "SELECT COUNT(" & strTablePrefix & "TOPICS.T_LAST_POST) AS NUM_ACTIVE "
			strSql = strSql & "FROM " & strTablePrefix & "TOPICS "
			strSql = strSql & "WHERE (((" & strTablePrefix & "TOPICS.T_LAST_POST)>'"& Session(strCookieURL & "last_here_date") & "'))"

			set rs = my_Conn.Execute(strSql)
			if not rs.EOF then
				ActiveTopicCount = rs("NUM_ACTIVE")
			else
				ActiveTopicCount = 0
			end if
		end if
	end if

	rs.close
	set rs = nothing
	ShowLastHere = (cint(ChkUser2(strDBNTUserName, Request.Cookies(strUniqueID & "User")("Pword"))) > 0)
%>
      <tr>
        <td bgcolor="<% =strCategoryCellColor %>" colspan="<% if (strShowModerators = "1") or (mlev = 4 or mlev = 3) then Response.Write("8") else Response.Write("6") end if %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" size="+1"><b>状态列(任意收缩)</b></font></td>
      </tr>
      <tr>
        <td rowspan="<% if ShowLastHere then Response.Write("5") else Response.Write("4") end if %>" bgcolor="<%= strForumCellColor %>">&nbsp;</td>
<%
	if ShowLastHere then 
%>
        <td bgcolor="<%= strForumCellColor %>" colspan="<% if ((strShowModerators = "1") or (mlev = "4" or mlev = "3")) then Response.Write("7") else Response.Write("5") end if %>">
        <font align=left face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">上次登陆时间：<% =ChkDate(Session(strCookieURL & "last_here_date")) %> <% =ChkTime(Session(strCookieURL & "last_here_date")) %></font>
        </td>
	  </tr>
	  <tr>
<%
	end if 
	if intPostCount > 0 then 
%>
        <td bgcolor="<%= strForumCellColor %>" colspan="<% if ((strShowModerators = "1") or (mlev = "4" or mlev = "3")) then Response.Write("7") else Response.Write("5") end if%>">
        <font face="<%= strDefaultFontFace %>" size="<% =strFooterFontSize %>"><% if Member_Count = 1 then Response.Write("1 Member has ") else Response.Write(Member_Count & " Members have ") end if %> made <% if intPostCount = 1 then Response.Write("1 post ") else Response.Write(intPostCount & " posts") end if %> in <% if intForumCount = 1 then Response.Write("1 forum") else Response.Write(intForumCount & " forums") end if %>
<%
		if (LastPostDate = "" or LastPostLink = "" or intPostCount = 0) then 
			Response.Write("") 
		else
			Response.Write(", with the last post on <a href=""" & lastPostLink & """>"& lastPostDate & "</a>")
			if  LastPostAuthorLink <> "" then
				Response.Write(LastPostAuthorLink & ".")
			else
				Response.Write(".")
			end if
		end if
%>
		  </font>
          </td>
        </tr>
        <tr>
<%
	end if
%>      
        <td bgcolor="<%= strForumCellColor %>" colspan="<% if ((strShowModerators = "1") or (mlev = "4" or mlev = "3")) then Response.Write("7") else Response.Write("5") end if%>">
        <font face="<%= strDefaultFontFace %>" size="<% =strFooterFontSize %>">There <% if intTopicCount = 1 then Response.Write("is ") else Response.Write("are ") end if%> currently <%= intTopicCount %> <%if intTopicCount = 1 then Response.Write(" topic ") else Response.Write(" topics") end if %><% if ActiveTopicCount > 0 then %> and <%= ActiveTopicCount %> <a href="active.asp">active <% if ActiveTopicCount = 1 then Response.Write("topic") else Response.Write("topics") end if %></a> since you last visited.<% elseif blnHiddenForums and (strLastPostDate > Session(strCookieURL & "last_here_date")) and ShowLastHere then %> and there are <a href="active.asp">active topics</a> since you last visited.<% elseif not(ShowLastHere) then Response.Write "." else %>and no active topics since you last visited.<% end if %></font>
        </td>
      </tr>
<%
	if NewMember_Name <> "" then 
%>
      <tr>          
        <td bgcolor="<%= strForumCellColor %>" colspan="<% if ((strShowModerators = "1") or (mlev = "4" or mlev = "3")) then Response.Write("7") else Response.Write("5") end if%>">
        <font face="<%= strDefaultFontFace %>" size="<% =strFooterFontSize %>">热烈欢迎我们的新会员：
<%
		if strUseExtendedProfile then
			Response.Write "        <a href=""pop_profile.asp?mode=display&id="& NewMember_Id & """>"
		else
			Response.Write "        <a href=""JavaScript:openWindow2('pop_profile.asp?mode=display&id=" & NewMember_Id & "')"">"
		end if
		Response.Write NewMember_Name & "</a>.</font>" & vbcrlf
%>
          </td>
        </tr>
<%
	end if 
end sub 
%>