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
<% set rs = Server.CreateObject("ADODB.Recordset") %>
<table width="100%" border="0">
  <tr>
    <td align="left" valign="top">
    <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    <img src="<%=strImageURL %>icon_folder_open.gif" border="0">&nbsp;<a href="default.asp">返回论坛首页</a>
<% If Request.QueryString("mode") = "DoIt" then %>
    &nbsp;<img src="<%=strImageURL %>icon_folder_open.gif" border="0">&nbsp;<a href="search.asp?Search=<%=Request.Form("Search")%>">论坛搜索</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" border="0">&nbsp;搜索结果：<%=Request.Form("Search")%>
<% else %>
    &nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" border="0">&nbsp;论坛搜索
<% end if %>
    </font>
    </td>
  </tr>
</table>
<%
If Request.QueryString("mode") = "DoIt" then
	if Request.Form("Search") <> "" or Request.Form("SearchMember") <> 0 then 'Huwr
		keywords = split(Request.Form("Search"), " ")
		keycnt = ubound(keywords)

		'## Forum_SQL - Find all records with the search criteria in them
		strSql = "SELECT " & strTablePrefix & "FORUM.FORUM_ID, " & strTablePrefix & "FORUM.F_SUBJECT, " & strTablePrefix & "FORUM.CAT_ID, " & strTablePrefix & "TOPICS.TOPIC_ID, " & strTablePrefix & "TOPICS.T_SUBJECT, " & strTablePrefix & "TOPICS.T_MAIL, " & strTablePrefix & "TOPICS.T_STATUS, " & strTablePrefix & "TOPICS.T_LAST_POST, " & strTablePrefix & "TOPICS.T_REPLIES, " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME "
		strSql = strSql & " FROM ((" & strTablePrefix & "FORUM LEFT JOIN " & strTablePrefix & "TOPICS "
		strSql = strSql & " ON " & strTablePrefix & "FORUM.FORUM_ID = " & strTablePrefix & "TOPICS.FORUM_ID) LEFT JOIN " & strTablePrefix & "REPLY "
		strSql = strSql & " ON " & strTablePrefix & "TOPICS.TOPIC_ID = " & strTablePrefix & "REPLY.TOPIC_ID) LEFT JOIN " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " ON " & strTablePrefix & "TOPICS.T_AUTHOR = " & strMemberTablePrefix & "MEMBERS.MEMBER_ID "
		strSql = strSql & " WHERE ("
	if Request.Form("Search") <> ""  then 'Huwr
	'################# New Search Code #################################################
			if Request.Form("SearchMessage") = 1 then
				if Request.Form("andor") = "phrase" then
					strSql = strSql & "     (" & strTablePrefix & "FORUM.F_SUBJECT LIKE '%" & ChkString(Request.Form("Search"), "SQLString") & "%'"
					strSql = strSql & " OR   " & strTablePrefix & "TOPICS.T_SUBJECT LIKE '%" & ChkString(Request.Form("Search"), "SQLString") & "%') "
				else
					For Each word in keywords
						SearchWord = ChkString(word, "SQLString")
						strSql = strSql & "     (" & strTablePrefix & "FORUM.F_SUBJECT LIKE '%" & SearchWord & "%'"
						strSql = strSql & " OR   " & strTablePrefix & "TOPICS.T_SUBJECT LIKE '%" & SearchWord & "%') "
			if cnt < keycnt then strSql = strSql &  Request.Form("andor") 
			cnt = cnt + 1
		next
				end if
			else
				if Request.Form("andor") = "phrase" then
						strSql = strSql & "     (" & strTablePrefix & "REPLY.R_MESSAGE LIKE '%" & ChkString(Request.Form("Search"), "SQLString") & "%'"
						strSql = strSql & " OR   " & strTablePrefix & "FORUM.F_DESCRIPTION LIKE '%" & ChkString(Request.Form("Search"), "SQLString") & "%'"
						strSql = strSql & " OR   " & strTablePrefix & "TOPICS.T_SUBJECT LIKE '%" & ChkString(Request.Form("Search"), "SQLString") & "%'"
						strSql = strSql & " OR   " & strTablePrefix & "TOPICS.T_MESSAGE LIKE '%" & ChkString(Request.Form("Search"), "SQLString") & "%') "
				else
					For Each word in keywords
						SearchWord = ChkString(word, "SQLString")
						strSql = strSql & "     (" & strTablePrefix & "REPLY.R_MESSAGE LIKE '%" & SearchWord & "%'"
						strSql = strSql & " OR   " & strTablePrefix & "FORUM.F_DESCRIPTION LIKE '%" & SearchWord & "%'"
						strSql = strSql & " OR   " & strTablePrefix & "TOPICS.T_SUBJECT LIKE '%" & SearchWord & "%'"
						strSql = strSql & " OR   " & strTablePrefix & "TOPICS.T_MESSAGE LIKE '%" & SearchWord & "%') "
						if cnt < keycnt then strSql = strSql &  Request.Form("andor") 
						cnt = cnt + 1
					next
				end if
			end if
			strSql = strSql & " ) "
		cnt = 0
'################# New Search Code #################################################
		if Request.Form("Forum") <> 0 then
			strSql = strSql & " AND " & strTablePrefix & "FORUM.FORUM_ID = " & Request.Form("Forum") & " "
		end if
		if Request.Form("SearchDate") <> 0 then
			dt = cint(Request.Form("SearchDate"))
			strSql = strSql & " AND (T_DATE > '" & DateToStr(dateadd("d", -dt, strForumTimeAdjust)) & "')"
		end if
		
		end if ' Huwr
		if Request.Form("Search") = "" and Request.Form("SearchMember") <> 0 then 'Huwr
				strSql = strSql & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & Request.Form("SearchMember") & " "
				strSql = strSql & " OR " & strTablePrefix & "REPLY.R_AUTHOR = " & Request.Form("SearchMember") & ") "
		else
			if Request.Form("SearchMember") <> 0 then
				strSql = strSql & " AND (" & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & Request.Form("SearchMember") & " "
				strSql = strSql & " OR " & strTablePrefix & "REPLY.R_AUTHOR = " & Request.Form("SearchMember") & ") "
			end if
		end if
		
		strSql = strSql & " AND " & strTablePrefix & "FORUM.F_TYPE = " & 0 
		strSql = strSql & " ORDER BY " & strTablePrefix & "TOPICS.FORUM_ID DESC, "
		strSql = strSql & "          " & strTablePrefix & "TOPICS.T_LAST_POST DESC"

		mypage = request("whichpage")

		If mypage = "" then
			mypage = 1
		end if

		rs.Open strSql, my_Conn, 3,1
%>
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td bgcolor="<% =strTableBorderColor %>">
    <table border="0" width="100%" cellspacing="1" cellpadding="4">
      <tr>
        <td align="center" bgcolor="<% =strHeadCellColor %>">&nbsp;</td>
        <td align="center" bgcolor="<% =strHeadCellColor %>"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">主题</font></b></td>
        <td align="center" bgcolor="<% =strHeadCellColor %>"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">作者</font></b></td>
        <td align="center" bgcolor="<% =strHeadCellColor %>"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">回覆数</font></b></td>
        <td align="center" bgcolor="<% =strHeadCellColor %>"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">最新文章</font></b></td>
      </tr>
<%		if rs.EOF or rs.BOF then  '## No new topic posts found in DB %>
      <tr>
        <td bgcolor="<% =strForumCellColor %>" colspan="5"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>没有找到相关资料</b></font></td>
      </tr>
<%
		else 
			rs.MoveFirst
			currForum = 0 
			currTopic = 0
			do until rs.EOF
				if ChkForumAccess(rs("FORUM_ID")) then 

					'## Forum_SQL - Find out if the Category is Locked or Un-Locked and if it Exists
					strSql = "SELECT " & strTablePrefix & "CATEGORY.CAT_STATUS "
					strSql = strSql & ", " & strTablePrefix & "FORUM.F_STATUS "
					strSql = strSql & " FROM " & strTablePrefix & "CATEGORY "
					strSql = strSql & " , " & strTablePrefix & "FORUM "
					strSql = strSql & " WHERE " & strTablePrefix & "CATEGORY.CAT_ID = " & strTablePrefix & "FORUM.CAT_ID "
					strSql = strSql & " AND " & strTablePrefix & "FORUM.FORUM_ID = " & rs("FORUM_ID")
		
					set rsCFStatus = my_Conn.Execute (StrSql)

					if (currForum <> rs("FORUM_ID")) and (currTopic <> rs("TOPIC_ID")) then 
%>
      <tr>
        <td height="20" colspan="5" bgcolor="<% =strCategoryCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>"><b>&nbsp;<% =ChkString(rs("F_SUBJECT"),"display") %></b></font></td>
      </tr>
						
<%					currForum = rs("FORUM_ID")
					end if %>
<%					if currTopic <> rs("TOPIC_ID") then %>
      <tr>
<%						if rsCFStatus("CAT_STATUS") <> 0 and rsCFstatus("F_STATUS") <> 0 and rs("T_STATUS") <> 0 then %>
        						<td bgcolor="<% =strForumCellColor %>" align="center"><a href="topic.asp?TOPIC_ID=<% =rs("TOPIC_ID") %>&FORUM_ID=<% =rs("FORUM_ID") %>&CAT_ID=<% =rs("CAT_ID") %>&Topic_Title=<% =ChkString(left(rs("T_SUBJECT"), 50),"urlpath") %>&Forum_Title=<% =ChkString(rs("F_SUBJECT"),"urlpath") %>"><% =ChkIsNew(rs("T_LAST_POST"),"Topic") %></a></td>
<%						else %>
						        <td bgcolor="<% =strForumCellColor %>" align="center"><a href="topic.asp?TOPIC_ID=<% =rs("TOPIC_ID") %>&FORUM_ID=<% =rs("FORUM_ID") %>&CAT_ID=<%=rs("CAT_ID") %>&Topic_Title=<% =ChkString(left(rs("T_SUBJECT"), 50),"urlpath") %>&Forum_Title=<%=ChkString(rs("F_SUBJECT"),"urlpath") %>"><img src="<%=strImageURL %>icon_folder_locked.gif"
<% 							if rsCFStatus("CAT_STATUS") = 0 then 
								Response.Write ("alt='分类已被锁定'")
							elseif rsCFStatus("F_STATUS") = 0 then 
								Response.Write ("alt='论坛已被锁定'")
							else
								Response.Write ("alt='主题已被锁定'")
							end if %>
								border="0"></a>
							</td>
<%						end if %>
        <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="topic.asp?TOPIC_ID=<% =rs("TOPIC_ID") %>&FORUM_ID=<% =rs("FORUM_ID") %>&CAT_ID=<% =rs("CAT_ID") %>&Topic_Title=<% =ChkString(left(rs("T_SUBJECT"), 50),"urlpath") %>&Forum_Title=<% =ChkString(rs("F_SUBJECT"),"urlpath") %>"><% =ChkString(left(rs("T_SUBJECT"), 50),"display") %></a>&nbsp;</font></td>
        <td bgcolor="<% =strForumCellColor %>" valign="top" align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strForumFontColor %>"><% =ChkString(rs("M_NAME"),"display") %></font></td>
        <td bgcolor="<% =strForumCellColor %>" valign="top" align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strForumFontColor %>"><% =rs("T_REPLIES") %></font></td>
        <td bgcolor="<% =strForumCellColor %>" valign="top" align="center" nowrap><font face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>" size="<% =strFooterFontSize %>"><b><% =ChkDate(rs("T_LAST_POST")) %></b>&nbsp;<% =ChkTime(rs("T_LAST_POST")) %></font></td>
      </tr>
<%					currTopic = rs("TOPIC_ID")
					end if 
				end if
				rs.MoveNext 
			loop 
		end if 
%>
    </table>
    </td>
  </tr>
</table>
<p align="center"><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>>
<a href="JavaScript:history.go(-1)">回到论坛搜索</a>
</font></p>

<%	Else %>

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">你必须输入关键字</p>

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript:history.go(-1)">回到论坛搜索</a></p>
<meta http-equiv="Refresh" content="0; URL=JavaScript:history.go(-1)">
<%
	end if
Else
%>

<form action="Search.asp?mode=DoIt" method="post">
<table border="0" cellspacing="0" cellpadding="0" align=center>
  <tr>
    <td bgcolor="<% =strPopUpBorderColor %>">
    <table border="0" cellspacing="1" cellpadding="5">
      <tr>
        <td bgColor="<% =strPopUpTableColor %>" align="right" valign="top"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">搜索目标：</font></td>
        <td bgColor="<% =strPopUpTableColor %>" align="left" valign="top"><input type="text" name="Search" size="40" value="<%=Request.QueryString("Search")%>"><br>
        <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
<% '################# New Search Code ################################################# %>
		<input type="radio" class="radio" name="andor" value="phrase">完全匹配<br>
<% '################# New Search Code ################################################# %>
		<input type="radio" class="radio" name="andor" value=" and " checked>符合所有字符串<br>
        <input type="radio" class="radio" name="andor" value=" or ">符合任一个字符串</font></td>
      </tr>
      <tr>
        <td bgColor="<% =strPopUpTableColor %>" align="right" valign="top"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">目标论坛：</font></td>
        <td bgColor="<% =strPopUpTableColor %>" align="left" valign="top">
        <select name="Forum" size="1">
          <option value="0">所有论坛</option>
<%
	'## Forum_SQL
	strSql = "SELECT " & strTablePrefix & "FORUM.FORUM_ID, " & strTablePrefix & "FORUM.F_SUBJECT "
	strSql = strSql & " FROM " & strTablePrefix & "FORUM"

	set rs = my_Conn.execute (strSql)

	On Error Resume Next

	do until rs.EOF
		if ChkForumAccess(rs("FORUM_ID")) then
			Response.Write "          <option value=""" & rs("FORUM_ID") & """>" & ChkString(left(rs("F_SUBJECT"), 50),"display") & "</option>" & vbCrLf
		end if
		rs.movenext
	loop
%>
        </select>
        </td>
      </tr>
<%'################# New Search Code #################################################%>
	  <tr>
        <td bgColor="<% =strPopUpTableColor %>" align="right" valign="top"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">搜索范围：</font></td>
        <td bgColor="<% =strPopUpTableColor %>" align="left" valign="top">
        <select NAME="SearchMessage">
          <option value="0">整篇文章</option>
          <option value="1">只有标题</option>		  
        </select>
        </td>
      </tr>
<%'################# New Search Code #################################################%>

      <tr>
        <td bgColor="<% =strPopUpTableColor %>" align="right" valign="top"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">按日期搜索：</font></td>
        <td bgColor="<% =strPopUpTableColor %>" align="left" valign="top">
        <select NAME="SearchDate">
          <option value="0">任何日期</option>
          <option VALUE="1">最近一天</option>
          <option VALUE="2">最近两天</option>
          <option VALUE="5">最近五天</option>
          <option VALUE="7">最近一个星期</option>
          <option VALUE="14">最近两个星期</option>
          <option VALUE="30">最近一个月前</option>
          <option VALUE="60">最近两个月</option>
          <option VALUE="120">最近四个月</option>
          <option VALUE="365">最近一年</option>
        </select>
        </td>
      </tr>
      <tr>
        <td bgColor="<% =strPopUpTableColor %>" align="right" valign="top"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">按用户名搜索：</font></td>
        <td bgColor="<% =strPopUpTableColor %>" align="left" valign="top">
        <select name="SearchMember">
          <option value="0">所有会员
<%
	'## Forum_SQL
	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS"
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_STATUS = " & 1
	strSql = strSql & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_NAME ASC;"
	
	set rs = my_Conn.execute (strSql)
	
	On Error Resume Next
	
	do until rs.EOF
		Response.Write "          <option value=""" & rs("MEMBER_ID") & """>" & ChkString(left(rs("M_NAME"), 50),"display") & "</option>" & vbCrLf
		rs.movenext
	loop
%>
        </select>
        </td>
      </tr>
      <tr>
        <td bgColor="<% =strPopUpTableColor %>" align="center" valign="top" colspan="2"><input type="submit" value="搜索"></td>
      </tr>
    </table>
    </td>
  </tr>
</table>

</form>
<% end if %>
<% on error resume next %>
<% set rs = nothing %>
<!--#INCLUDE FILE="inc_footer.asp" -->
