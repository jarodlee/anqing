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
<!--#INCLUDE FILE="inc_functions.asp" -->
<!--#INCLUDE FILE="inc_top_short.asp" -->
<% 
'## Forum_SQL
strSql = "SELECT " & strTablePrefix & "ANNOUNCE.A_ID" 
strSql = strSql & ", " & strTablePrefix & "ANNOUNCE.A_AUTHOR"
strSql = strSql & ", " & strTablePrefix & "ANNOUNCE.A_SUBJECT"
strSql = strSql & ", " & strTablePrefix & "ANNOUNCE.A_MESSAGE"
'strSql = strSql & ", " & strTablePrefix & "ANNOUNCE.A_FORUMLIST"
strSql = strSql & ", " & strTablePrefix & "ANNOUNCE.A_START_DATE"
strSql = strSql & ", " & strTablePrefix & "ANNOUNCE.A_END_DATE"
strSql = strSql & " FROM " & strTablePrefix & "ANNOUNCE "
if mLev < 4 then
	strSql = strSql & " WHERE " & strTablePrefix & "ANNOUNCE.A_START_DATE <= " & "'" & DatetoStr(Now()) & "'"
	strSql = strSql & " AND " & strTablePrefix & "ANNOUNCE.A_END_DATE > " & "'" & DatetoStr(Now()) & "'"
end if

set rs = my_Conn.Execute (StrSql)

if rs.EOF or rs.BOF then
	Response.Redirect(strForumURL)
else

if (mLev = 4) or (lcase(strNoCookies) = "1") then
 	AdminAllowed = 1
else
 	AdminAllowed = 0
end if
%>
<br>
<table border="0" width="100%" cellspacing="1" cellpadding="4" align="center" bgcolor="<% =strTableBorderColor %>">
      <tr>
        <td align="center" bgcolor="<% =strHeadCellColor %>" width="100%"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">��̳������</font></b></td>
      </tr>
<%
	'## Forum_SQL
	strSql = "SELECT " & strTablePrefix & "ANNOUNCE.A_ID"
	strSql = strSql & ", " & strTablePrefix & "ANNOUNCE.A_AUTHOR"
	strSql = strSql & ", " & strTablePrefix & "ANNOUNCE.A_SUBJECT"
	strSql = strSql & ", " & strTablePrefix & "ANNOUNCE.A_MESSAGE"
	strSql = strSql & ", " & strTablePrefix & "ANNOUNCE.A_START_DATE"
	strSql = strSql & ", " & strTablePrefix & "ANNOUNCE.A_END_DATE"
	strSql = strSql & " FROM " & strTablePrefix & "ANNOUNCE "
	if mLev < 4 then
		strSql = strSql & " WHERE " & strTablePrefix & "ANNOUNCE.A_START_DATE <= " & "'" & DatetoStr(Now()) & "'"
		strSql = strSql & " AND " & strTablePrefix & "ANNOUNCE.A_END_DATE > " & "'" & DatetoStr(Now()) & "'"
        end if
	strSql = strSql & " ORDER BY " & strTablePrefix & "ANNOUNCE.A_START_DATE DESC"
	strSql = strSql & ", " & strTablePrefix & "ANNOUNCE.A_ID DESC;"


	set rs = Server.CreateObject("ADODB.Recordset")
	rs.cachesize = 20
	rs.open  strSql, my_Conn, 3

	if not(rs.EOF or rs.BOF) then  '## Replies found in DB
		rs.movefirst
		rs.pagesize = strPageSize
		maxpages = cint(rs.pagecount)
	end if
	if rs.EOF or rs.BOF then  '## No replies found in DB
		Response.Write ""
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

        <tr bgcolor="<% =CColor %>" valign="top" nowrap>
		<td ><font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b><% =ChkString(rs("A_SUBJECT"),"display") %></b></font></td></tr>
        <tr bgcolor="<% =CColor %>" valign="top"><td ><img src="<%=strImageURL %>icon_posticon.gif" border="0" hspace="3"><font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">�����ˣ�<% =rs("A_AUTHOR") %><br>(<% =ChkDate(rs("A_START_DATE")) %>&nbsp;��&nbsp;<% =ChkDate(rs("A_END_DATE")) %>)<% if AdminAllowed = 1 then %>&nbsp;<a href="JavaScript:openWindow('pop_announce_delete.asp?mode=Announcement&A_ID=<% =rs("A_ID") %>')"><img src="<%=strImageURL %>icon_delete_reply.gif" height=15 width=15 alt="ɾ������" border="0" align="absmiddle" hspace="6"></a><% end if %><% if rs("A_END_DATE") <  DatetoStr(Now()) then %><font color="<% =strActiveLinkColor %>">�����򹫸��ѹ��ڣ�<% end if %><% if rs("A_START_DATE") > DatetoStr(Now()) then %><font color="<% =strActiveLinkColor %>">�����򹫸���δ������<% end if %></font></font>

        <hr noshade size="<% =strFooterFontSize %>">
        
        <font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% =formatStr(rs("A_MESSAGE")) %><a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" height=15 width=15 border="0" align="right" alt="�ص�ҳ��"></a></font></td></tr>

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
    </table>
<br>
<!--#INCLUDE FILE="inc_footer_short.asp" -->
<%
end if
%>