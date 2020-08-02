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
select case Request.QueryString("mode") 
	case "CloseTopic"
		mLev = cint(ChkUser2(Request.Form("user"), Request.Form("Pass"))) 
		if mLev > 0 then  '## is Member
			if (chkForumModerator(Request.Form("FORUM_ID"), Request.Form("user")) = "1") _
			or (mLev = 4) _
			or (chkForumModerator(Request.Form("FORUM_ID"), Session(strCookieURL & "userid")) = "1") then

				'## Forum_SQL
				strSql = "Update " & strTablePrefix & "TOPICS "
				strSql = strSql & " SET " & strTablePrefix & "TOPICS.T_STATUS = 0 "
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
<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">主题已被关闭！</font></p>
<%			Else %>
<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">你没权限关闭主题！</font><br>
<br>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript: onClick= history.go(-1)">返回重新认证</a></font></p>
<%			end if %>
<%		Else %>
<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">你没权限关闭主题！</font><br>
<br>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript: onClick= history.go(-1)">返回重新认证</a></font></p>
<%
		end if 

	case "CloseForum"

		mLev = cint(ChkUser2(Request.Form("user"), Request.Form("Pass"))) 
		if mLev > 0 then  '## is Member
			if (chkForumModerator(Request.Form("FORUM_ID"), Request.Form("user")) = "1") or (mLev = 4) or (chkForumModerator(Request.Form("FORUM_ID"), Session(strCookieURL & "userid")) = "1") then

				'## Forum_SQL
				strSql = "Update " & strTablePrefix & "FORUM "
				strSql = strSql & " SET " & strTablePrefix & "FORUM.F_STATUS = 0 "
				strSql = strSql & " WHERE " & strTablePrefix & "FORUM.FORUM_ID = " & Request.Form("FORUM_ID") 

				my_Conn.Execute (strSql)
%>
<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">论坛已被锁定！</font></p>
<%			else %>
<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><b>你没权限关闭论坛！</font><br>
<br>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript: onClick= history.go(-1)">返回重新认证</a></font></p>
<%			end if %>
<%		else %>
<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><b>你没权限关闭论坛！</font><br>
<br>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript: onClick= history.go(-1)">返回重新认证</a></font></p>
<%
		end if 

	case "CloseCategory"

		mLev = cint(ChkUser(Request.Form("user"), Request.Form("Pass"))) 
		if mLev > 0 then  '## is Member
			if mLev = 4 then

				'## Forum_SQL
				strSql = "Update " & strTablePrefix & "CATEGORY "
				strSql = strSql & " SET " & strTablePrefix & "CATEGORY.CAT_STATUS = 0 "
				strSql = strSql & " WHERE " & strTablePrefix & "CATEGORY.CAT_ID = " & Request.Form("CAT_ID") 

				my_Conn.Execute (strSql)
%>
<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><b>分类已被锁定!</b></font></p>
<%			else %>
<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><b>你没权限关闭分类！</font><br>
<br>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript: onClick= history.go(-1)">返回重新认证</a></font></p>
<%			end if %>
<%		else %>
<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><b>你没权限关闭分类！</font><br>
<br>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript: onClick= history.go(-1)">返回重新认证</a></font></p>
<%
		end if 

	case "LockMember"

		mLev = cint(ChkUser(Request.Form("user"), Request.Form("Pass"))) 
		if mLev > 0 then  '## is Member
			if mLev = 4 then

				'## Forum_SQL
				strSql = "Update " & strMemberTablePrefix & "MEMBERS "
				strSql = strSql & " SET " & strMemberTablePrefix & "MEMBERS.M_STATUS = 0 "
				strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & Request.Form("MEMBER_ID") 

				my_Conn.Execute (strSql)
%>
<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><b>会员已被锁定!</b></font></p>
<%			else %>
<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><b>你没权限关锁定会员！</font><br>
<br>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript: onClick= history.go(-1)">返回重新认证</a></font></p>
<%			end if %>
<%		else %>
<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><b>你没权限关锁定会员！</font><br>
<br>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript: onClick= history.go(-1)">返回重新认证</a></font></p>
<%
		end if 
	case else 
%>
<P><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">锁定 <% if Request.Querystring("mode") = "Topic" then Response.Write("主题") %><% if Request.Querystring("mode") = "Forum" then Response.Write("论坛") %><% if Request.Querystring("mode") = "Category" then Response.Write("分类") %><% if Request.Querystring("mode") = "Member" then Response.Write("会员") %></font></p>

<p><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b><font color=red>注意：</font></b>  
<%		select case Request.QueryString("mode") %>
<%			case "Member" %>
只有版主跟管理员才能锁定一个会员.
<%			case "Category" %>
只有版主跟管理员才能锁定一个分类.
<%			case "Forum" %>
只有版主跟管理员才能锁定一个论坛.
<%			case "Topic" %>
只有版主跟管理员才能锁定一个主题.
<%		end select %>
</font></p>

<form action="pop_lock.asp?mode=<% if Request.Querystring("mode") = "Topic" then Response.Write("CloseTopic") %><% if Request.Querystring("mode") = "Forum" then Response.Write("CloseForum") %><% if Request.Querystring("mode") = "Category" then Response.Write("CloseCategory") %><% if Request.Querystring("mode") = "Member" then Response.Write("LockMember") %>" method=post>
<input type=hidden name="Method_Type" value="<% if Request.Querystring("mode") = "Topic" then Response.Write("CloseTopic") %><% if Request.Querystring("mode") = "Forum" then Response.Write("CloseForum") %><% if Request.Querystring("mode") = "Category" then Response.Write("CloseCategory") %><% if Request.Querystring("mode") = "Member" then Response.Write("LockMember") %>">
<input type=hidden name="TOPIC_ID" value="<% =Request.QueryString("TOPIC_ID") %>">
<input type=hidden name="FORUM_ID" value="<% =Request.QueryString("Forum_ID") %>">
<input type=hidden name="CAT_ID" value="<% =Request.QueryString("CAT_ID") %>">
<input type=hidden name="MEMBER_ID" value="<% =Request.QueryString("MEMBER_ID") %>">

<table border="0" width="75%" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="<% =strPopUpBorderColor %>">
    <table border="0" width="100%" cellspacing="1" cellpadding="1">
<%		if strAuthType="db" then %>
      <tr>
        <td bgcolor=<% =strPopUpTableColor %> align=right nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">你的名字：</font></b></td>
        <td bgcolor=<% =strPopUpTableColor %>><input type=text name="User" value="<% =Request.Cookies(strUniqueID & "User")("Name")%>" size=20></td>
      </tr>
      <tr>
        <td bgcolor=<% =strPopUpTableColor %> align=right nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">你的密码：</FONT></b></td>
        <td bgcolor=<% =strPopUpTableColor %>><input type=Password name="Pass" value="<% =Request.Cookies(strUniqueID & "User")("Pword")%>" size=20></td>
      </tr>
<%		else %>
<%			if strAuthType="nt" then %>
      <tr>
        <td bgcolor=<% =strPopUpTableColor %> align=right nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">NT 帐号：</font></b></td>
        <td bgcolor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><%=Session(strCookieURL & "userid")%></FONT></td>
      </tr>
<%			end if %>
<%		end if %>
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
