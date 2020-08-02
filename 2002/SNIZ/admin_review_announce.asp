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
<% If Session(strCookieURL & "Approval") = "15916941253" Then %>
<!--#INCLUDE file="inc_functions.asp" -->
<!--#INCLUDE file="inc_top.asp" -->
<table border="0" width="100%" align="center">
  <tr>
    <td width="33%" align="left" nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    <img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="default.asp">返回论坛首页</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="admin_home.asp">论坛管理中心</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="admin_announce_home.asp">论坛公告中心</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" height=15 width=15 border="0">&nbsp;查看修改公告</font></td>
  </tr>
</table>
<br>
<table border="0" width="80%" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td bgcolor="<% =strTableBorderColor %>">
    <table width="100%" align="center" border="0" cellspacing="1" cellpadding="3">
      <tr>
        <td align="center" bgcolor="<% =strHeadCellColor %>" colspan="4"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">查看修改公告</font></td>
      </tr>
      <tr>
        <td align="center" bgcolor="<% =strCategoryCellColor %>"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">公告<font size="<% =strFooterFontSize %>">（点选公告编辑）</font></font></b></td>
        <td align="center" bgcolor="<% =strCategoryCellColor %>"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">开始日期</font></b></td>
        <td align="center" bgcolor="<% =strCategoryCellColor %>"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">终止日期</font></b></td>
        <td align="center" bgcolor="<% =strCategoryCellColor %>"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>"></font></b></td>
      </tr>
<% 
	'## Forum_SQL
	strSql = "SELECT " & strTablePrefix & "ANNOUNCE.A_ID" 
	strSql = strSql & ", " & strTablePrefix & "ANNOUNCE.A_AUTHOR"
	strSql = strSql & ", " & strTablePrefix & "ANNOUNCE.A_SUBJECT"
	strSql = strSql & ", " & strTablePrefix & "ANNOUNCE.A_START_DATE"
	strSql = strSql & ", " & strTablePrefix & "ANNOUNCE.A_END_DATE"
	strSql = strSql & " FROM " & strTablePrefix & "ANNOUNCE "
	strSql = strSql & " ORDER BY " & strTablePrefix & "ANNOUNCE.A_START_DATE ASC;"


	set rs = Server.CreateObject("ADODB.Recordset")
	rs.cachesize = 20
	rs.open  strSql, my_Conn, 3

	if not(rs.EOF or rs.BOF) then  '## Replies found in DB
		rs.movefirst
		rs.pagesize = strPageSize
		maxpages = cint(rs.pagecount)
	end if
	if rs.EOF or rs.BOF then  '## No replies found in DB
%>
      <tr>
        <td bgcolor="<% =strForumFirstCellColor %>" colspan="4"><font face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>" size="<% =strDefaultFontSize %>" valign="top"><b>暂时没有任何公告</b></font></td>
      </tr>
<%
	else
		'rs.movefirst
		intI = 0 
		howmanyrecs = 0
		rec = 1

		do until rs.EOF '**
			if intI = 0 then
				CColor = strForumFirstCellColor
			else
				CColor = strAltForumCellColor
			end if
 %>
      <tr>
        <td bgcolor="<% =CColor %>" valign="top" align="center" nowrap>
		<font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="admin_edit_announce.asp?A_ID=<% =rs("A_ID") %>&A_SUBJECT=<% =ChkString(rs("A_SUBJECT"),"urlpath") %>"><% =ChkString(rs("A_SUBJECT"),"display") %></a></font></td>
        <td bgcolor="<% =CColor %>" valign="top" align="center">
        	<font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% =ChkDate(rs("A_START_DATE")) %></font></td>
        <td bgcolor="<% =CColor %>" valign="top" align="center">
        	<font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% =ChkDate(rs("A_END_DATE")) %></font></td>
        <td bgcolor="<% =CColor %>" valign="top" align="center">
        	<font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript:openWindow('pop_announce_delete.asp?mode=Announcement&A_ID=<% =rs("A_ID") %>')"><img src="<%=strImageURL %>icon_trashcan.gif" alt="删除公告" border="0" hspace="0"></a></font></td>
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
 %>
    </table></td>
  </tr>
</table>
<br>
<!--#INCLUDE file="inc_footer.asp" -->
<% else %>
<% Response.Redirect "admin_login.asp" %>
<% end if %>