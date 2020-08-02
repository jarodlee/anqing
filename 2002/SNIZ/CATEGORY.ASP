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

'## Forum_SQL - Get all Forums From DB
strSql = "SELECT " & strTablePrefix & "CATEGORY.CAT_ID, " & strTablePrefix & "CATEGORY.CAT_STATUS, " 
strSql = strSql & strTablePrefix & "CATEGORY.CAT_NAME "
strSql = strSql & " FROM " & strTablePrefix & "CATEGORY "
strSql = strSql & " WHERE ("  & strTablePrefix & "CATEGORY.CAT_ID = " & intQS_CATID & ")"
strSql = strSql & " ORDER BY " & strTablePrefix & "CATEGORY.CAT_NAME ASC;"

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
        <td align=center bgcolor="<% =strHeadCellColor %>" nowrap valign="top"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">论坛列表</font></b></td>
        <td align=center bgcolor="<% =strHeadCellColor %>" nowrap valign="top"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">主题</font></b></td>
        <td align=center bgcolor="<% =strHeadCellColor %>" nowrap valign="top"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">文章</font></b></td>
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
        <td bgcolor="<% =strCategoryCellColor %>" colspan="<%=sGetColspan(6, 5)%>"><font face="<% =strDefaultFontFace %>" color="<% =strCategoryFontColor %>" size="<% =strDefaultFontSize %>" valign="top"><b>没有找到分类或论坛</b></font></td>
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

		'## Forum_SQL - Build SQL to get forums via category
		strSql = "SELECT " & strTablePrefix & "FORUM.FORUM_ID, " 
		strSql = strSql & strTablePrefix & "FORUM.F_STATUS, " 
		strSql = strSql & strTablePrefix & "FORUM.CAT_ID, " 
		strSql = strSql & strTablePrefix & "FORUM.F_SUBJECT, " 
		strSql = strSql & strTablePrefix & "FORUM.F_URL, " 
		strSql = strSql & strTablePrefix & "FORUM.F_DESCRIPTION, " 
		strSql = strSql & strTablePrefix & "FORUM.CAT_ID, " 
		strSql = strSql & strTablePrefix & "FORUM.F_TOPICS, " 
		strSql = strSql & strTablePrefix & "FORUM.F_COUNT, " 
		strSql = strSql & strTablePrefix & "FORUM.F_LAST_POST, " 
		strSql = strSql & strTablePrefix & "FORUM.F_STATUS, " 
		strSql = strSql & strTablePrefix & "FORUM.F_TYPE,  " 
		strSql = strSql & strTablePrefix & "FORUM.F_PRIVATEFORUMS,  " 
		strSql = strSql & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " 
		strSql = strSql & strMemberTablePrefix & "MEMBERS.M_NAME " 
		strSql = strSql & "FROM " & strTablePrefix & "FORUM " 
		strSql = strSql & "LEFT JOIN " & strMemberTablePrefix & "MEMBERS ON "
		strSql = strSql & strTablePrefix & "FORUM.F_LAST_POST_AUTHOR = "
		strSql = strSql & strMemberTablePrefix & "MEMBERS.MEMBER_ID "
		strSql = strSql & " WHERE " & strTablePrefix & "FORUM.CAT_ID = " & rs("CAT_ID") & " "
		strSql = strSql & " ORDER BY " & strTablePrefix & "FORUM.F_SUBJECT ASC;"

		set rsForum =  my_Conn.Execute (strSql)
				
		chkDisplayHeader = true
				 
		if rsForum.eof or rsForum.bof then
%>

      <tr>
        <td bgcolor="<% =strCategoryCellColor %>" colspan="<%=sGetColspan(6, 5)%>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" valign="top"><b><% =ChkString(rs("CAT_NAME"),"display") %></b></font></td>
<%
			if (mlev = 4 or mlev = 3) or (lcase(strNoCookies) = "1") then 
%>
        <td bgcolor="<% =strCategoryCellColor %>" align=center valign=top nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% call CategoryAdminOptions() %></font></b></td>
<%
			end if 
%>
      </tr>
<%
			Response.Write	"  <tr>" & vbCrLf & _
							"    <td bgcolor=""" & strForumCellColor & """ colspan=""" & sGetColspan(6, 5)
			Response.Write """><font face=""" & strDefaultFontFace & """ color=""" & strForumFontColor & """ size=""" & strDefaultFontSize & """><b>没有任何论坛</b></font></td>" & vbCrLf
			if (mlev = 4 or mlev = 3) then
				Response.Write "    <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ color=""" & strForumFontColor & """ size=""" & strDefaultFontSize & """>&nbsp;</font></td>" & vbCrLf
			end if
			Response.Write "  </tr>" & vbCrLf
		else
			blnHiddenForums = false
			do until rsForum.Eof
				if ChkDisplayForum(rsForum("FORUM_ID")) then
					if rsForum("F_TYPE") <> "1" then 
						intPostCount  = intPostCount + rsForum("F_COUNT")
						intTopicCount = intTopicCount + rsForum("F_TOPICS")
						intForumCount = intForumCount + 1
						if rsForum("F_LAST_POST") > strLastPostDate then 
							strLastPostDate = rsForum("F_LAST_POST")
							intLastPostForum_ID = rsForum("FORUM_ID")
						end if
					end if
					if chkDisplayHeader then
						Response.Write	"      <tr>" & vbcrlf & _
										"        <td bgcolor=""" & strCategoryCellColor & """ colspan=""" & sGetColspan(6, 5)
						Response.Write """ valign=top><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """><b>" & ChkString(rs("CAT_NAME"),"display") & "</b></font></td>" & vbcrlf 
						if (mlev = 4 or mlev = 3) or (lcase(strNoCookies) = "1") then 
							Response.Write "        <td bgcolor=""" & strCategoryCellColor & """ align=center valign=top nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>"
							call CategoryAdminOptions()
							Response.Write "</font></b></td>" & vbcrlf 
						end if 
						Response.Write "      </tr>" & vbcrlf
						chkDisplayHeader = false
					end if
					Response.Write	"      <tr>" & vbcrlf & _
									"        <td bgcolor=""" & strForumCellColor & """ align=center valign=top>" & vbcrlf
					if rsForum("F_TYPE") = 0 then
						if rs("CAT_STATUS") = 0 then 
							Response.Write "        <a href=""forum.asp?FORUM_ID=" & rsForum("FORUM_ID") & "&CAT_ID=" & rsForum("CAT_ID") & "&Forum_Title=" & ChkString(rsForum("F_SUBJECT"),"urlpath") & """>"
							if rsForum("F_LAST_POST") > Session(strCookieURL & "last_here_date") then
								Response.Write "<img src=""<%=strImageURL %>icon_folder_new_locked.gif"" height=15 width=15 border=0 hspace=0 alt=""分类已锁定""></a>"
							else
								Response.Write "<img src=""<%=strImageURL %>icon_folder_locked.gif"" height=15 width=15 border=0 hspace=0 alt=""分类已锁定""></a>"
							end if     
						else 
							if rsForum("F_STATUS") <> 0 then 
								Response.Write "        <a href=""forum.asp?FORUM_ID=" & rsForum("FORUM_ID") & "&CAT_ID=" & rsForum("CAT_ID") & "&Forum_Title=" & ChkString(rsForum("F_SUBJECT"),"urlpath") & """>" & ChkIsNew(rsForum("F_LAST_POST"),"") & "</a>"
							else
								Response.Write "        <a href=""forum.asp?FORUM_ID=" & rsForum("FORUM_ID") & "&CAT_ID=" & rsForum("CAT_ID") & "&Forum_Title=" & ChkString(rsForum("F_SUBJECT"),"urlpath") & """>"
								if rsForum("F_LAST_POST") > Session(strCookieURL & "last_here_date") then
									Response.Write "<img src=""<%=strImageURL %>icon_folder_new_locked.gif"" height=15 width=15 border=0 hspace=0 alt=""讨论区已锁定""></a>"
								else
									Response.Write "<img src=""<%=strImageURL %>icon_folder_locked.gif"" height=15 width=15 border=0 hspace=0 alt=""讨论区已锁定""></a>"
								end if
							end if
						end if
					else 
						if  rsForum("F_TYPE") = 1 then 
							Response.Write "<a href=""" & rsForum("F_URL") & """ target=""_blank""><img src=""<%=strImageURL %>icon_url.gif"" height=16 width=16 border=0 hspace=0></a>"
						end if
					end if 
					Response.Write	"        </td>" & vbcrlf
%>
        <td<% if rsForum("F_TYPE") = 1 then Response.Write(" colspan=4") %> bgcolor="<% =strForumCellColor %>" valign="top">
        <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="<% if rsForum("F_TYPE") = 0 then Response.Write("forum.asp?FORUM_ID=" & rsForum("FORUM_ID") & "&CAT_ID=" & rsForum("CAT_ID") & "&Forum_Title=" & ChkString(rsForum("F_SUBJECT"),"urlpath")) else if rsForum("F_TYPE") = 1 then Response.Write(rsForum("F_URL") & """ target=""_blank") end if %>"><% =ChkString(rsForum("F_SUBJECT"),"display") %></a><br>
        <font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
        <% = rsForum("F_DESCRIPTION") %>
        </font></td>
<%
					if rsForum("F_TYPE") = 0 then
						if IsNull(rsForum("F_TOPICS")) then 
%>
        <td bgcolor="<% =strForumCellColor %>" align=center valign=top><font face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>" size="<% =strDefaultFontSize %>">0</font></td>
<%
						else 
%>
        <td bgcolor="<% =strForumCellColor %>" align=center valign=top><font face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>" size="<% =strDefaultFontSize %>"><% =rsForum("F_TOPICS") %></font></td>
<%
						end if 
						if IsNull(rsForum("F_COUNT")) then 
%>
        <td bgcolor="<% =strForumCellColor %>" align=center valign=top><font face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>" size="<% =strDefaultFontSize %>">0</font></td>
<%
						else 
%>
        <td bgcolor="<% =strForumCellColor %>" align=center valign=top><font face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>" size="<% =strDefaultFontSize %>"><% =rsForum("F_COUNT") %></font></td>
<%
						end if 
						if IsNull(rsForum("MEMBER_ID")) then
							strLastUser = ""
						else
							strLastUser = "<br>作者：" 
							if strUseExtendedProfile then
								strLastUser = strLastUser & "<a href=""pop_profile.asp?mode=display&id="& RsForum("MEMBER_ID") & """>"
							else
								strLastUser = strLastUser & "<a href=""JavaScript:openWindow2('pop_profile.asp?mode=display&id=" & rsForum("MEMBER_ID") & "')"">"
							end if
							strLastUser = strLastUser & ChkString(rsForum("M_NAME"),"display") & "</a>"
						end if
%>
        <td bgcolor="<% =strForumCellColor %>" align=center valign=top nowrap><font face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>" size="<% =strFooterFontSize %>"><b><% =ChkDate(rsForum("F_LAST_POST")) %></b><br>
        <% =ChkTime(rsForum("F_LAST_POST")) %><%= strLastUser%></font></td>
<%
					else 
						if rsForum("F_TYPE") = 1 then 
							'## Do Nothing 
						end if 
					end if 
					if (strShowModerators = "1") or (mlev = 4 or mlev = 3) then 
%>
        <td bgcolor="<% =strForumCellColor %>" align=left valign=top><font face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>" size="<% =strDefaultFontSize %>">
<%
						if (listForumModerators(rsForum("FORUM_ID")) <> "") then
							Response.Write(listForumModerators(rsForum("FORUM_ID")))
						else
							Response.Write("&nbsp;")
						end if 
%>
        </font></td>
<%
					end if 
					if (mlev = 4 or mlev = 3) or (lcase(strNoCookies) = "1") then 
%>
        <td bgcolor="<% =strForumCellColor %>" align=center valign=top nowrap><% call ForumAdminOptions %></td>
<%
					end if 
%>
      </tr>
<%
				else
					blnHiddenForums = true
				end if ' ChkDisplayForum() 
				rsForum.MoveNext
			loop
		end if
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
set rsForum = nothing 
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
          <a href="JavaScript:openWindow('pop_lock.asp?mode=Forum&FORUM_ID=<% =rsForum("FORUM_ID") %>&CAT_ID=<% =rsForum("CAT_ID") %>&Forum_Title=<% =ChkString(rsForum("F_SUBJECT"),"JSurlpath")%>')"><img src="<%=strImageURL %>icon_lock.gif" alt="锁定讨论区" border="0" hspace="0"></a>
<%
				else 
%>
          <a href="JavaScript:openWindow('pop_open.asp?mode=Forum&FORUM_ID=<% =rsForum("FORUM_ID") %>&CAT_ID=<% =rsForum("CAT_ID") %>&Forum_Title=<% =ChkString(rsForum("F_SUBJECT"),"JSurlpath")%>')"><img src="<%=strImageURL %>icon_unlock.gif" alt="解开讨论区锁定" border="0" hspace="0"></a>
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
          <a href="JavaScript:openWindow('pop_delete.asp?mode=Forum&FORUM_ID=<% =rsForum("FORUM_ID") %>&CAT_ID=<% =rsForum("CAT_ID") %>&Forum_Title=<% =ChkString(rsForum("F_SUBJECT"),"JSurlpath") %>')"><img src="<%=strImageURL %>icon_trashcan.gif" alt="删除论坛" border="0" hspace="0"></a>
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
%>