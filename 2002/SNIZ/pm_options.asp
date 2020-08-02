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

<% Response.Buffer = True
if Request.Cookies(strUniqueID & "User")("Name") = "" Then %>
<meta http-equiv="Refresh" content="0; URL=<% =Request.ServerVariables("HTTP_REFERER") %>">
<center>
<br><br><br><br>
<font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><B>登陆论坛才能使用这个功能</b></font>
<br>
<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="<% =Request.ServerVariables("HTTP_REFERER") %>">返回论坛</font></a></p>
<br><br><br><br><br>
</center>
<% else 
if Request.QueryString("mode") = "setoptions" then
'## Forum_SQL
strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
strSql = strSql & " SET M_PMRECEIVE = " & Request.Form("statusstorage") & ", "
strSql = strSql & "     M_PMEMAIL = " & Request.Form("emailstorage")
strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & Request.Cookies(strUniqueID & "User")("Name") & "'"
strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_PASSWORD = '" & Request.Cookies(strUniqueID & "User")("PWord") & "'"

my_Conn.Execute(strSql)
if strSetCookieToForum = 1 then
   Response.Cookies("paging").Path = strCookieURL
end if
   Response.Cookies("paging")("outbox") = Request.Form("layoutstorage")
   Response.Cookies("paging").Expires = dateAdd("d", 360, strForumTimeAdjust)
	if request.cookies("paging")("outbox") = "" then request.cookies("paging")("outbox") = "single"

%>
<center>
<br>
<br>
<br>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>悄悄话设置已更新</b></font>
<br>

<br>
<br>
<br>
<br>
<br>

<% else 

		'## Forum_SQL
		strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.M_PASSWORD, " & strMemberTablePrefix & "MEMBERS.M_PMRECEIVE, " & strMemberTablePrefix & "MEMBERS.M_PMEMAIL "
		strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & Request.Cookies(strUniqueID & "User")("Name") & "'"
		strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_PASSWORD = '" & Request.Cookies(strUniqueID & "User")("PWord") & "'"

		set rs = my_Conn.Execute(strSql)

%>
<center>
<table width=100% border=0 cellspacing=1 cellpadding=4 bgcolor="<% =strTableBorderColor %>">
<form action="pm_options.asp?mode=setoptions" method="POST">
<TR bgcolor="<% =strHeadCellColor %>">
<TD><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>"><B>启动/关闭 悄悄话信息讯息</B></FONT></TD>
</tr>
<tr bgcolor="<% =strForumFirstCellColor %>">
  <td>
<font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
<b><% =strForumTitle %> 悄悄话信息目前状态为：<% if rs("M_PMRECEIVE") = "1" then %>启动<% else %>关闭<% end if %></b>。<br>你可以从底下选项将它<% if rs("M_PMRECEIVE") = "1" then %>关闭<% else %>启动<% end if %>。关闭后你将再收不到悄悄话讯息<br>
<INPUT TYPE="RADIO" NAME="statusstorage" VALUE="1" <% if rs("M_PMRECEIVE") = "1" then Response.Write("checked")%>> 启动悄悄话讯息。 
<INPUT TYPE="RADIO" NAME="statusstorage" VALUE="0" <% if rs("M_PMRECEIVE") = "0" then Response.Write("checked")%>> 关闭悄悄话讯息。<br>
<font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">你可以随时回到此页切换悄悄话讯息的状态</font>
</FONT>
</TD></TR>
<TR bgcolor="<% =strHeadCellColor %>">
  <TD><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>"><B>电子邮件通知</B></FONT>
    </TD>
</tr>
<tr bgcolor="<% =strForumFirstCellColor %>"><td><font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b><% =strForumTitle %> 可以在当你收到悄悄话讯息时，发送一封电子邮件通知</b><br>
<INPUT TYPE="RADIO" NAME="emailstorage" VALUE="1" <% if rs("M_PMEMAIL") = "1" then Response.Write("checked")%>> 希望收到电子邮件通知。 
<INPUT TYPE="RADIO" NAME="emailstorage" VALUE="0" <% if rs("M_PMEMAIL") = "0" then Response.Write("checked")%>> 不要收到通知讯息。</FONT>
</TD>
</TR>
<TR bgcolor="<% =strHeadCellColor %>">
<TD><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>"><B>收件箱/发件箱 设定</B></FONT></TD>
</tr>
<tr bgcolor="<% =strForumFirstCellColor %>">
<td>
<font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
<b>Please select your layout preference.</b><br>
<INPUT TYPE="RADIO" NAME="layoutstorage" VALUE="single" <% if request.cookies("paging")("outbox") = "single" then Response.Write("checked")%>> 单页布局：&nbsp;<font size="1">收件箱和发件箱出现在同一个页面里。</font><br>
<INPUT TYPE="RADIO" NAME="layoutstorage" VALUE="double" <% if request.cookies("paging")("outbox") = "double" then Response.Write("checked")%>> 双页布局：&nbsp;&nbsp;<font size="1">收件箱和发件箱分别出现在两个页面里。</font><br>
<INPUT TYPE="RADIO" NAME="layoutstorage" VALUE="none" <% if request.cookies("paging")("outbox") = "none" then Response.Write("checked")%>> 没有收件箱：&nbsp;<font size="1">没有收件箱.  发送的信息是不被存储的。</font><br>
</FONT>
</Td></Tr>

</TABLE>
<P>
<INPUT TYPE="submit" VALUE="更新设置">
<% end if %>
<% end if %>
</P>
<!--#INCLUDE FILE="inc_footer_short.asp" -->