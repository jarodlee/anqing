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
<!--#INCLUDE FILE="inc_top.asp" -->
<%

if Request.Cookies(strUniqueID & "User")("Name") = "" Then %>
<center>
<br><br><br><br>
<font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><B>ֻ�е�½����̳����ʹ�����Ļ�����</b></font>
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
    <img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="default.asp">������̳��ҳ</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="pm_view.asp">���Ļ���Ϣ</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" height=15 width=15 border="0">&nbsp;�������Ļ�<br>
    </font></td>
  </tr>
</table>
<table width=100% border=0 cellspacing=1 cellpadding=4 bgcolor="<% =strTableBorderColor %>">
<TR>
<td align="center" bgcolor="<% =strHeadCellColor %>" width="<% =strTopicWidthLeft %>" <% if lcase(strTopicNoWrapLeft) = "1" then Response.Write(" nowrap") %>><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">����</font></b></td>
<td align="left" bgcolor="<% =strHeadCellColor %>" width="<% =strTopicWidthRight %>" <% if lcase(strTopicNoWrapRight) = "1" then Response.Write(" nowrap") %>><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">���⣺&nbsp;&nbsp; <% =rsMessage("M_SUBJECT") %></font></b></td>
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
        <br><font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">��������<% =rs("M_POSTS") %></font>
        </td>
<TD bgcolor="<% =strForumFirstCellColor %>" valign="top">
<img src="<%=strImageURL %>icon_posticon.gif" border="0" hspace="3"><font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">&nbsp;<% =ChkDate(rsMessage("M_SENT")) %>&nbsp;&nbsp;<% =ChkTime(rsMessage("M_SENT")) %></font>
        <% if strUseExtendedProfile then %>
		&nbsp;<a href="pop_profile.asp?mode=display&id=<% =rs("MEMBER_ID") %>"><img src="<%=strImageURL %>1-profile.gif" alt="�鿴��Ա����" border="0" align="absmiddle" hspace="6"></a>
        <% else %>
		&nbsp;<a href="JavaScript:openWindow3('pop_profile.asp?mode=display&id=<% =rs("MEMBER_ID") %>')"><img src="<%=strImageURL %>1-profile.gif" alt="�鿴��Ա����" border="0" align="absmiddle" hspace="6"></a>
		<% end if %>	
        &nbsp;<a href="JavaScript:openWindow('pop_mail.asp?id=<% =rs("M_FROM") %>')"><img src="<%=strImageURL %>1-email.gif" alt="�����ʼ�" border="0" align="absmiddle" hspace="6"></a>
<%			if strHomepage = "1" then %>
<%				if rs("M_Homepage") <> " " then %>
        &nbsp;<a href="<% =rs("M_Homepage") %>"><img src="<%=strImageURL %>1-home.gif" alt="��� <% =rs("M_NAME") %> ����ҳ" border="0" align="absmiddle" hspace="6"></a>
<%				end if %>
<%			end if %>
			
<%			if strICQ = "1" then %>
<%			  if rs("M_ICQ") <> " " then %>
        &nbsp;<a href="JavaScript:openWindowICQ('pop_messengers.asp?mode=ICQ&ICQ=<% =rs("M_ICQ") %>&M_NAME=<% =ChkString(rs("M_NAME"),"urlpath") %>')"><img src="http://online.mirabilis.com/scripts/online.dll?icq=<% =rs("M_ICQ") %>&img=5" height=15 width=15 alt="����ICQѶϢ�� <% =rs("M_NAME") %>" border="0" align="absmiddle" hspace="6"></a>
<%			  end if %>
<%			end if %>
<%			if strYAHOO = "1" then %>
<%			  if rs("M_YAHOO") <> " " then %>
        &nbsp;<a href="http://search.tencent.com/cgi-bin/friend/user_show_info?ln=<% =rs("M_YAHOO") %>"  target=_blank><img align=absmiddle width=16 height=16 src="http://icon.tencent.com/<% =rs("M_YAHOO") %>/s/00/" alt="<% =rs("M_YAHOO") %>" border=0></a> <% =rs("M_YAHOO") %>
<%			  end if %>
<%			end if %>
<%			if (strAIM = "1") then %>
<%				if rs("M_AIM") <> " " then %>
        &nbsp;<a href="JavaScript:openWindow('pop_messengers.asp?mode=AIM&AIM=<% =rs("M_AIM") %>&M_NAME=<% =ChkString(rs("M_NAME"),"urlpath") %>')"><img src="<%=strImageURL %>icon_aim.gif" height=15 width=15 alt="��AIMѶϢ�� <% =rs("M_NAME") %>" border="0" align="absmiddle" hspace="6"></a>
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