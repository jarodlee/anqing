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
%>
<!--#INCLUDE FILE="config.asp" -->
<% If Session(strCookieURL & "Approval") = "15916941253" Then %>
<!--#INCLUDE file="inc_functions.asp" -->
<!--#INCLUDE file="inc_top.asp" -->
<table border="0" width="100%" align="center">
  <tr>
    <td width="33%" align="left" nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    <img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="default.asp">������̳��ҳ</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="admin_home.asp">��̳��������</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="admin_announce_home.asp">��̳��������</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" height=15 width=15 border="0">&nbsp;�鿴�޸Ĺ���</font></td>
  </tr>
</table>
<br>
<table border="0" width="80%" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td bgcolor="<% =strTableBorderColor %>">
    <table width="100%" align="center" border="0" cellspacing="1" cellpadding="3">
      <tr>
        <td align="center" bgcolor="<% =strHeadCellColor %>" colspan="4"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">�鿴�޸Ĺ���</font></td>
      </tr>
      <tr>
        <td align="center" bgcolor="<% =strCategoryCellColor %>"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">����<font size="<% =strFooterFontSize %>">����ѡ����༭��</font></font></b></td>
        <td align="center" bgcolor="<% =strCategoryCellColor %>"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">��ʼ����</font></b></td>
        <td align="center" bgcolor="<% =strCategoryCellColor %>"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">��ֹ����</font></b></td>
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
        <td bgcolor="<% =strForumFirstCellColor %>" colspan="4"><font face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>" size="<% =strDefaultFontSize %>" valign="top"><b>��ʱû���κι���</b></font></td>
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
        	<font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript:openWindow('pop_announce_delete.asp?mode=Announcement&A_ID=<% =rs("A_ID") %>')"><img src="<%=strImageURL %>icon_trashcan.gif" alt="ɾ������" border="0" hspace="0"></a></font></td>
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