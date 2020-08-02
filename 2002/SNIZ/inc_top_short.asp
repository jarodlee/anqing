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

	strDBNTUserName = Request.Cookies(strUniqueID & "User")("Name")
	strDBNTFUserName = Request.Form("Name")
	if strAuthType = "nt" then
		strDBNTUserName = Session(strCookieURL & "userID")
		strDBNTFUserName = Session(strCookieURL & "userID")
	end if

set my_Conn= Server.CreateObject("ADODB.Connection")
my_Conn.Open strConnString

%>

<html>

<head>
<title><% =strForumTitle %></title>
<meta name="copyright" content="Snitz Forums 2000 Version 3.1 SR4 资源搜罗站汉化修改(www.99ss.net)">
<style type=text/css><!--
a:link    {color:<% =strLinkColor %>;text-decoration:<% =strLinkTextDecoration %>}
a:visited {color:<% =strVisitedLinkColor %>;text-decoration:<% =strVisitedTextDecoration %>}
a:hover   {color:<% =strHoverFontColor %>;text-decoration:<% =strHoverTextDecoration %>}
input.radio {background: <% = strPopUpTableColor %>; color:#000000}
font {  font-size: 9pt; line-height: 13pt; FONT-FAMILY:<% =strDefaultFontFace %>}
td {  font-size: 9pt; line-height: 13pt; FONT-FAMILY:<% =strDefaultFontFace %>}
textarea { BACKGROUND-COLOR: #e8e8e8; BORDER-BOTTOM: 1px double; BORDER-LEFT: 1px double; BORDER-RIGHT: 1px double; BORDER-TOP: 1px double; COLOR: #000000;; font-size: 9pt ;FONT-FAMILY:<% =strDefaultFontFace %>}
input { BACKGROUND-COLOR: #e8e8e8; BORDER-BOTTOM: 1px double; BORDER-LEFT: 1px double; BORDER-RIGHT: 1px double; BORDER-TOP: 1px double; COLOR: #000000;; font-size: 9pt; FONT-FAMILY:<% =strDefaultFontFace %>}

--></style>
</head>
  
<BODY background="<%= strPageBGImage %>" bgColor="<% =strPageBGColor %>" text="<% =strDefaultFontColor %>" link="<% =strLinkColor %>" aLink=<% =strActiveLinkColor %> vLink="<% =strActiveLinkColor %>" onLoad="window.focus()">

<table width="100%" height="100%">
  <tr>
    <td align=center valign=center>
    <div align=center><center>
    <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
