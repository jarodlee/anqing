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

if Request.Cookies(strUniqueID & "User")("Name") = "" Then %>
<center>
<br><br><br><br>
<font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><B>只有登陆了论坛才能使用悄悄话功能</b></font>
<br><br><br><br><br><br>
</center>
<% else
     
	
	strSql = "SELECT "   & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME,  " & strTablePrefix & "PM.M_ID,  " & strTablePrefix & "PM.M_TO, " & strTablePrefix & "PM.M_SUBJECT, " & strTablePrefix & "PM.M_MESSAGE, " & strTablePrefix & "PM.M_SENT, " & strTablePrefix & "PM.M_FROM "
	strSql = strSql & " FROM " & strTablePrefix & "MEMBERS , " & strTablePrefix & "PM "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & Request.Cookies(strUniqueID & "User")("Name") & "'"
	strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "PM.M_FROM "
	strSql = strSql & " AND " & strMemberTablePrefix & "PM.M_ID =  " & Request.QueryString("id")
	strSql = strSql & " ORDER BY " & strTablePrefix & "PM.M_SENT DESC" 
	
	Set rsMessage = my_Conn.Execute(strSql)
	
	if rsMessage.BOF or rsMessage.EOF then
	   Response.Redirect("pm_view.asp")
	end if
	
	strSql ="SELECT " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_ICQ, " & strMemberTablePrefix & "MEMBERS.M_YAHOO, " & strMemberTablePrefix & "MEMBERS.M_AIM, " & strMemberTablePrefix & "MEMBERS.M_TITLE, " & strMemberTablePrefix & "MEMBERS.M_TITLE, " & strMemberTablePrefix & "MEMBERS.M_Homepage, " & strMemberTablePrefix & "MEMBERS.M_LEVEL, " & strMemberTablePrefix & "MEMBERS.M_POSTS, " & strMemberTablePrefix & "MEMBERS.M_CITY, " & strMemberTablePrefix & "MEMBERS.M_STATE, " & strMemberTablePrefix & "MEMBERS.M_COUNTRY, " & strTablePrefix & "PM.M_FROM, " & strTablePrefix & "PM.M_SUBJECT "    
	strSql = strSql & " FROM " & strTablePrefix & "MEMBERS , " & strTablePrefix & "PM "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "PM.M_FROM "
	strSql = strSql & " AND " & strMemberTablePrefix & "PM.M_ID =  " & Request.QueryString("id")
	
	
	Set rs = my_Conn.Execute(strSql)
	if strI = 0 then 
				CColor = strAltForumCellColor
			else
				CColor = strForumCellColor
			end if
	
%>
<center>
<table border="0" width="100%" align="center">
  <tr>
    <td width="33%" align="left" nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    <img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="default.asp">返回论坛首页</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="pm_view.asp">悄悄话信息</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" height=15 width=15 border="0">&nbsp;备份悄悄话<br>
    </font></td>
  </tr>
</table>
<table width=100% border=0 cellspacing=1 cellpadding=4 bgcolor="<% =strTableBorderColor %>">
<TR>
<td align="center" bgcolor="<% =strHeadCellColor %>" width="<% =strTopicWidthLeft %>" <% if lcase(strTopicNoWrapLeft) = "1" then Response.Write(" nowrap") %>><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">作者</font></b></td>
<td align="left" bgcolor="<% =strHeadCellColor %>" width="<% =strTopicWidthRight %>" <% if lcase(strTopicNoWrapRight) = "1" then Response.Write(" nowrap") %>><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">主题：&nbsp;&nbsp; <% =rsMessage("M_SUBJECT") %></font></b></td>
</TR>
<tr>
        <td bgcolor="<% =strForumFirstCellColor %>" valign="top">
        <font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
        <% if strUseExtendedProfile then %>
		<a href="pop_profile.asp?mode=display&id=<% =rs("M_FROM") %>">
        <% else %>
		<a href="JavaScript:openWindow3('pop_profile.asp?mode=display&id=<% =rs("M_FROM") %>')">
		<% end if %>	
		<font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b><% =ChkString(rs("M_NAME"),"display") %></a>
        </b></font>
<%		if strShowRank = 1 or strShowRank = 3 then %>
        <br><font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>"><% = ChkString(getMember_Level(rs("M_TITLE"), rs("M_LEVEL"), rs("M_POSTS")),"display") %></font>
<%		end if %>
<%		if strShowRank = 2 or strShowRank = 3 then %>
        <br><% = getStar_Level(rs("M_LEVEL"), rs("M_POSTS")) %>
<%		end if %>
        <br>
        <br><font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>"><% =rs("M_COUNTRY") %></font>
        <br><font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">发表数：<% =rs("M_POSTS") %></font>
        </td>
<TD bgcolor="<% =strForumFirstCellColor %>" valign="top">
<img src="<%=strImageURL %>icon_posticon.gif" border="0" hspace="3"><font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">&nbsp;<% =ChkDate(rsMessage("M_SENT")) %>&nbsp;&nbsp;<% =ChkTime(rsMessage("M_SENT")) %></font>
        <% if strUseExtendedProfile then %>
		&nbsp;<a href="pop_profile.asp?mode=display&id=<% =rs("MEMBER_ID") %>"><img src="<%=strImageURL %>1-profile.gif" alt="查看会员资料" border="0" align="absmiddle" hspace="6"></a>
        <% else %>
		&nbsp;<a href="JavaScript:openWindow3('pop_profile.asp?mode=display&id=<% =rs("MEMBER_ID") %>')"><img src="<%=strImageURL %>1-profile.gif" alt="查看会员资料" border="0" align="absmiddle" hspace="6"></a>
		<% end if %>	
        &nbsp;<a href="JavaScript:openWindow('pop_mail.asp?id=<% =rs("M_FROM") %>')"><img src="<%=strImageURL %>1-email.gif" alt="电子邮件" border="0" align="absmiddle" hspace="6"></a>
<%			if strHomepage = "1" then %>
<%				if rs("M_Homepage") <> " " then %>
        &nbsp;<a href="<% =rs("M_Homepage") %>"><img src="<%=strImageURL %>1-home.gif" alt="浏览 <% =rs("M_NAME") %> 的主页" border="0" align="absmiddle" hspace="6"></a>
<%				end if %>
<%			end if %>
			
<%			if strICQ = "1" then %>
<%			  if rs("M_ICQ") <> " " then %>
        &nbsp;<a href="JavaScript:openWindowICQ('pop_messengers.asp?mode=ICQ&ICQ=<% =rs("M_ICQ") %>&M_NAME=<% =ChkString(rs("M_NAME"),"urlpath") %>')"><img src="http://online.mirabilis.com/scripts/online.dll?icq=<% =rs("M_ICQ") %>&img=5" height=15 width=15 alt="发送ICQ讯息给 <% =rs("M_NAME") %>" border="0" align="absmiddle" hspace="6"></a>
<%			  end if %>
<%			end if %>
<%			if strYAHOO = "1" then %>
<%			  if rs("M_YAHOO") <> " " then %>
        &nbsp;<a href="http://search.tencent.com/cgi-bin/friend/user_show_info?ln=<% =rs("M_YAHOO") %>"  target=_blank><img align=absmiddle width=16 height=16 src="http://icon.tencent.com/<% =rs("M_YAHOO") %>/s/00/" alt="<% =rs("M_YAHOO") %>" border=0></a> <% =rs("M_YAHOO") %>
<%			  end if %>
<%			end if %>
<%			if (strAIM = "1") then %>
<%				if rs("M_AIM") <> " " then %>
        &nbsp;<a href="JavaScript:openWindow('pop_messengers.asp?mode=AIM&AIM=<% =rs("M_AIM") %>&M_NAME=<% =ChkString(rs("M_NAME"),"urlpath") %>')"><img src="<%=strImageURL %>icon_aim.gif" height=15 width=15 alt="传AIM讯息给 <% =rs("M_NAME") %>" border="0" align="absmiddle" hspace="6"></a>
<%				end if %>
<%			end if %>
<HR>
<font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% =formatStr(rsMessage("M_MESSAGE")) %></font>
</td>
</tr>


</TABLE>
</center>

<% end if %>
<br>
<!--#INCLUDE FILE="inc_footer.asp" -->