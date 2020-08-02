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
<!--#INCLUDE FILE="inc_top_short.asp" -->
<%
Select case Request.QueryString("mode") 
	case "OpenTopic"
		mLev = cint(ChkUser(Request.Form("user"), Request.Form("Pass"))) 
		if mLev > 0 then  '## is Member
			if (chkForumModerator(Request.Form("FORUM_ID"), Request.Form("User")) = "1") or _
			(mLev = 4) or _
			(chkForumModerator(Request.Form("FORUM_ID"), Session(strCookieURL & "userid")) = "1") then

				'## Forum_SQL
				strSql = "UPDATE " & strTablePrefix & "TOPICS "
				strSql = strSql & " SET " & strTablePrefix & "TOPICS.T_STATUS = 1 "
'Lock in place code#################################
				if Request.Form("InPlace") = "1" then
					strSQL = strSql & ", " & strTablePrefix & "TOPICS.T_INPLACE = 1"
				else
					strSQL = strSql & ", " & strTablePrefix & "TOPICS.T_INPLACE = 0"
				end if
				strSql = strSql & " WHERE " & strTablePrefix & "TOPICS.TOPIC_ID = " & Request.Form("TOPIC_ID") 
'Lock in place code#################################


				my_Conn.Execute (strSql)

%>
<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><b>主题锁定解除！</b></font></p>
<%			else %>
<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><b>你没有权限解除锁定的主题<br>
<br>
<a href="JavaScript: onClick= history.go(-1)">返回</a></font></p>
<%			end if %>
<%		else %>
<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><b>你没有权限解除锁定的主题<br>
<br>
<a href="JavaScript: onClick= history.go(-1)">返回</a></font></p>
<%
		end if 
	case "OpenForum"
		mLev = cint(ChkUser2(Request.Form("user"), Request.Form("Pass"))) 
		if mLev > 0 then  '## is Member
			if (chkForumModerator(Request.Form("FORUM_ID"), Request.Form("user")) = "1") or (mLev = 4) or (chkForumModerator(Request.Form("FORUM_ID"), Session(strCookieURL & "userid")) = "1")then

				'## Forum_SQL
				strSql = "UPDATE " & strTablePrefix & "FORUM "
				strSql = strSql & " SET " & strTablePrefix & "FORUM.F_STATUS = 1 "
				strSql = strSql & " WHERE " & strTablePrefix & "FORUM.FORUM_ID = " & Request.Form("FORUM_ID") 

				my_Conn.Execute (strSql)

%>
<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><b>论坛锁定解除！</b></font></p>
<%			else %>
<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><b>你没有权限解除锁定的论坛<br>
<br>
<a href="JavaScript: onClick= history.go(-1)">返回重新认证</a></font></p>
<%			end if %>
<%		else %>
<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><b>你没有权限解除锁定的论坛<br>
<br>
<a href="JavaScript: onClick= history.go(-1)">返回重新认证</a></font></p>
<%
		end if 
	case "OpenCategory"
		mLev = cint(ChkUser(Request.Form("user"), Request.Form("Pass"))) 
		if mLev > 0 then  '## is Member
			if mLev =4 then

				'## Forum_SQL
				strSql = "UPDATE " & strTablePrefix & "CATEGORY "
				strSql = strSql & " SET " & strTablePrefix & "CATEGORY.CAT_STATUS = 1 "
				strSql = strSql & " WHERE " & strTablePrefix & "CATEGORY.CAT_ID = " & Request.Form("CAT_ID") 

				my_Conn.Execute (strSql)
%>
<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><b>分类锁定解除！</b></font></p>
<%			else %>
<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><b>你没有权限解除锁定的分类<br>
<br>
<a href="JavaScript: onClick= history.go(-1)">返回重新认证</a></font></p>
<%			end if %>
<%		else %>
<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><b>你没有权限解除锁定的分类<br>
<br>
<a href="JavaScript: onClick= history.go(-1)">返回重新认证</a></font></p>
<%
		end if 
	case "UnLockMember"
		mLev = cint(ChkUser(Request.Form("user"), Request.Form("Pass"))) 
		if mLev > 0 then  '## is Member
			if mLev =4 then

				'## Forum_SQL
				strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
				strSql = strSql & " SET " & strMemberTablePrefix & "MEMBERS.M_STATUS = 1 "
				strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & Request.Form("MEMBER_ID") 

				my_Conn.Execute (strSql)
%>
<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><b>会员锁定解除！</b></font></p>
<%			else %>
<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><b>你没有权限解除锁定的会员<br>
<br>
<a href="JavaScript: onClick= history.go(-1)">返回重新认证</a></font></p>
<%			end if %>
<%		else %>
<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><b>你没有权限解除锁定的会员<br>
<br>
<a href="JavaScript: onClick= history.go(-1)">返回重新认证</a></font></p>
<%
		end if 
	case else 
%>
<P><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">解除<% if Request.Querystring("mode") = "Topic" then Response.Write("主题") %><% if Request.Querystring("mode") = "Forum" then Response.Write("论坛") %><% if Request.Querystring("mode") = "Category" then Response.Write("分类") %><% if Request.Querystring("mode") = "Member" then Response.Write("会员") %>锁定</font></p>

<p><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b><font color=red>注意：</font></b>  
<%				if Request.QueryString("mode") = "Member" then %>
只有论坛管理员可以解除会员锁定。
<%				else %>
<%					if Request.QueryString("mode") = "Category" then %>
只有论坛管理员可以解除分类锁定。
<%					else %>
<%						if Request.QueryString("mode") = "Forum" then %>
只有论坛管理员可以解除论坛锁定。
<%						else %>
<%							if Request.QueryString("mode") = "Topic" then %>
只有论坛管理员和版主可以解除主题锁定。
<%							end if %>
<%						end if %>
<%					end if %>
<%				end if %>
</font></p>

<form action="pop_open.asp?mode=<% if Request.Querystring("mode") = "Topic" then Response.Write("OpenTopic") %><% if Request.Querystring("mode") = "Forum" then Response.Write("OpenForum") %><% if Request.Querystring("mode") = "Category" then Response.Write("OpenCategory") %><% if Request.Querystring("mode") = "Member" then Response.Write("UnLockMember") %>" method=Post>
<input type=hidden name="TOPIC_ID" value="<% =Request.QueryString("TOPIC_ID") %>">
<input type=hidden name="FORUM_ID" value="<% =Request.QueryString("Forum_ID") %>">
<input type=hidden name="CAT_ID" value="<% =Request.QueryString("CAT_ID") %>">
<input type=hidden name="MEMBER_ID" value="<% =Request.QueryString("MEMBER_ID") %>">

<table border="0" width="75%" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="<% =strPopUpBorderColor %>">
    <table border="0" width="100%" cellspacing="1" cellpadding="1">
<%				if strAuthType="db" then %>
      <tr>
        <td bgColor=<% =strPopUpTableColor %> align=right nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">你的名称：</font></b></td>
        <td bgColor=<% =strPopUpTableColor %>><input type=Text name="User" value="<% =Request.Cookies(strUniqueID & "User")("Name")%>" size=20></td>
      </tr>
      <tr>
        <td bgColor=<% =strPopUpTableColor %> align=right nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">你的密码：</FONT></b></td>
        <td bgColor=<% =strPopUpTableColor %>><input type=Password name="Pass" value="<% =Request.Cookies(strUniqueID & "User")("Pword")%>" size=20></td>
      </tr>
<%				else %>
<%					if strAuthType="nt" then %>
      <tr>
        <td bgColor=<% =strPopUpTableColor %> align=right nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">NT 帐号：</font></b></td>
        <td bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><%=Session(strCookieURL & "userid")%></font></td>
      </tr>
<%					end if %>
<%				end if %>   
<%'Lock in place code#################################
	if Request.QueryString("mode") = "Topic" Then %> 
	<tr>
	<% strSQL = "SELECT " & strTablePrefix & "TOPICS.T_INPLACE FROM " & strTablePrefix & "TOPICS "
		strSql = strSQL & "WHERE " & strTablePrefix & "TOPICS.TOPIC_ID = " & Request.QueryString("TOPIC_ID") 
		set rs = my_conn.Execute(strSql)
		%>
		<td bgColor=<% =strPopUpTableColor %> align=right><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">锁於某处</font></td>
		<td bgColor=<% =strPopUpTableColor %> align=left><Input type="Checkbox" value="1" name="InPlace" <% if rs("T_INPLACE") = 1 then response.write "checked" %>></td>
	</tr> 
<% 	End If 
'Lock in place code#################################%> 			
      <tr>
        <td bgColor=<% =strPopUpTableColor %> colspan=2 align=center><Input type=Submit value="提交"></td>
      </tr>
    </table>
    </td>
  </tr>
</table>
</form>
<% end select %>
<!--#INCLUDE FILE="inc_footer_short.asp" -->
