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
<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">������ɾ����</font></p>
<P align=center>(<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>�ǵ���������ҳ�档</b></font>)</p>
<%			else %>
<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><b>����Ȩɾ������</font><br>
<br>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript: onClick= history.go(-1) ">����������֤</a></font></p>
<%			end if %>
<%		else %>
<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><b>����Ȩɾ������</font><br>
<br>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript: onClick= history.go(-1)">����������֤</a></font></p>
<%
		end if
else
%>
<P><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">ɾ������</font></p>

<p><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b><font color=red>ע�⣺</font></b>ֻ�й���Ա����ɾ�����棡</font></p>

<form action="pop_announce_delete.asp?mode=<%if Request.QueryString("mode") = "Announcement" then Response.Write("DeleteAnnouncement")%>" method=post id=Form1 name=Form1>
<input type=hidden name="A_ID" value="<% =Request.QueryString("A_ID") %>">
<table border="0" width="75%" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="<% =strPopUpBorderColor %>">
    <table border="0" width="100%" cellspacing="1" cellpadding="1">
<%				if strAuthType="db" then %>
      <tr>
        <td bgColor=<% =strPopUpTableColor %> align=right nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">������֣�</font></b></td>
        <td bgColor=<% =strPopUpTableColor %>><input type=text name="User" value="<% =Request.Cookies(strUniqueID & "User")("Name")%>" size=20></td>
      </tr>
      <tr>
        <td bgColor=<% =strPopUpTableColor %> align=right nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">������룺</FONT></b></td>
        <td bgColor=<% =strPopUpTableColor %>><input type=Password name="Pass" value="<% =Request.Cookies(strUniqueID & "User")("Pword")%>" size=20></td>
      </tr>
<%				else %>
<%					if strAuthType="nt" then %>
      <tr>
        <td bgColor=<% =strPopUpTableColor %> align=right nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">NT �ʺţ�</font></b></td>
        <td bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><%=Session(strCookieURL & "userid")%></font></td>
      </tr>
<%					end if %>
<%				end if %>
      <tr>
        <td bgColor=<% =strPopUpTableColor %> colspan=2 align=center><Input type=Submit value="�ύ" id=Submit1 name=Submit1></td>
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