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
if strAuthType = "db" then
	strDBNTUserName = Request.Form("User")
end if

if Request.QueryString("mode") = "DeleteAnnouncement" then
	mLev = cint(ChkUser2(strDBNTUserName, Request.Form("Pass")))
		if mLev > 0 then  '## is Member
			if mLev = 4 then
				'## Forum_SQL - Delete all replys related to the topics
				strSql = "DELETE FROM " & strTablePrefix & "ANNOUNCE "
				strSql = strSql & " WHERE " & strTablePrefix & "ANNOUNCE.A_ID = " & Request.Form("A_ID")

                		my_Conn.Execute strSql
%>
<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">公告已删除！</font></p>
<P align=center>(<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>记得重新整理页面。</b></font>)</p>
<%			else %>
<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><b>你无权删除公告</font><br>
<br>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript: onClick= history.go(-1) ">返回重新认证</a></font></p>
<%			end if %>
<%		else %>
<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><b>你无权删除公告</font><br>
<br>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript: onClick= history.go(-1)">返回重新认证</a></font></p>
<%
		end if
else
%>
<P><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">删除公告</font></p>

<p><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b><font color=red>注意：</font></b>只有管理员可以删除公告！</font></p>

<form action="pop_announce_delete.asp?mode=<%if Request.QueryString("mode") = "Announcement" then Response.Write("DeleteAnnouncement")%>" method=post id=Form1 name=Form1>
<input type=hidden name="A_ID" value="<% =Request.QueryString("A_ID") %>">
<table border="0" width="75%" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="<% =strPopUpBorderColor %>">
    <table border="0" width="100%" cellspacing="1" cellpadding="1">
<%				if strAuthType="db" then %>
      <tr>
        <td bgColor=<% =strPopUpTableColor %> align=right nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">你的名字：</font></b></td>
        <td bgColor=<% =strPopUpTableColor %>><input type=text name="User" value="<% =Request.Cookies(strUniqueID & "User")("Name")%>" size=20></td>
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
      <tr>
        <td bgColor=<% =strPopUpTableColor %> colspan=2 align=center><Input type=Submit value="提交" id=Submit1 name=Submit1></td>
      </tr>
    </table>
    </td>
  </tr>
</table>
</form>
<%
end if
%>
<!--#INCLUDE FILE="inc_footer_short.asp" -->