<%
'########## Snitz Forums 2000 Version 3.1 SR4 ####################
'#                                                               #
'#  �����޸�: ��Դ����վ                                         #
'#  �����ʼ�: cgier@21cn.com                                     #
'#  ��ҳ��ַ: http://www.sdsea.com                               #
'#            http://www.99ss.net                                #
'#            http://www.cdown.net                               #
'#	     http://www.wzdown.com                               #
'#	     http://www.13888.net                                #
'#  ��̳��ַ:http://ubb.yesky.net                                #
'#  ����޸�����: 2001/03/12    ���İ汾��Version 3.1 SR4        #
'#################################################################
'# ԭʼ��Դ                                                      #
'# Snitz Forums 2000 Version 3.1 SR4                             #
'# Copyright 2000 http://forum.snitz.com - All Rights Reserved   #
'#################################################################
'#����Ȩ������                                                   #
'#                                                               #
'# ������Ϊ��������(shareware)�ṩ������վ���ʹ�ã�����Ƿ��޸�,#
'# ת�أ�ɢ��������������ͼ����Ϊ��������ɾ����Ȩ������          #
'# ���������վ��ʽ����������ű�������֪ͨ���ǣ��Ա������ܹ�֪��#
'# ������ܣ�����������վ�������ǵ����ӣ�ϣ���ܸ��������лл��  #
'#################################################################
'# �����������ǵ��Ͷ��Ͱ�Ȩ����Ҫɾ�����ϵİ�Ȩ�������֣�лл����#
'# �����κ������뵽���ǵ���̳��������                            #
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
<meta name="copyright" content="Snitz Forums 2000 Version 3.1 SR4 ��Դ����վ�����޸�(www.99ss.net)">
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
