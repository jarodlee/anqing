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
<!--#INCLUDE FILE="inc_top_short.asp" -->
<% select case Request.QueryString("mode") %>
<%	case "ICQ" %>
<p><font face="<% =strDefaultFontFace %>"size="<% =strHeaderFontSize %>">����ICQѶϢ</font></p>
<form action="http://wwp.mirabilis.com/scripts/WWPMsg.dll" method="post">
<input type="hidden" name="subject" value="From Your Web Page">
<input type="hidden" name="to" value="<% =Request.QueryString("ICQ") %>">

<table border="0" width="75%" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="<% =strPopUpBorderColor %>">
    <table border="0" width="100%" cellspacing="1" cellpadding="1">
      <TR>
        <TD bgColor=<% =strPopUpTableColor %> align=right nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�ռ��ˣ�</font></td>
        <TD bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% =Request.QueryString("M_NAME")%></font></td>
      </tr>
      <TR>
        <TD bgColor=<% =strPopUpTableColor %> align=right nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�ռ��� ICQ ���룺</font></td>
        <TD bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><img src="http://online.mirabilis.com/scripts/online.dll?icq=<% =Request.QueryString("ICQ") %>&img=5" border="0" align="absmiddle"><% =Request.QueryString("ICQ") %></font></td>
      </tr>
      <TR>
        <TD bgColor=<% =strPopUpTableColor %> align=right nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">������������</font></td>
        <TD bgColor=<% =strPopUpTableColor %>><input type="text" name="from" value size="20" maxlength="40" onfocus="this.select()"></td>
      <TR>
        <TD bgColor=<% =strPopUpTableColor %> align=right nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�����˵����ʼ���</FONT></td>
        <TD bgColor=<% =strPopUpTableColor %>><input type="text" name="fromemail" value size="20" maxlength="40" onfocus="this.select()"></td>
      <TR>
        <TD bgColor=<% =strPopUpTableColor %> Colspan=2 align=center nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">ѶϢ</font></td>
      </TR>
      <TR>
        <TD bgColor=<% =strPopUpTableColor %> Colspan=2 align=center><textarea name="body" rows="3" cols="40" wrap="Virtual"></textarea></td>
      </TR>
      <TR>
        <TD bgColor=<% =strPopUpTableColor %> Colspan=2 align=center><input type=submit value="����"></td>
      </tr>
    </TABLE>
    </td>
  </tr>
</table>
</form>
<%	case "AIM" %>
<p><font face="<% =strDefaultFontFace %>"size="<% =strHeaderFontSize %>"><% =Request.QueryString("M_NAME") %>��AIM����</font></p>

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>ע�⣺</b> ����밲װ�� AOL Instant Messenger ����ִ�д˹��ܡ�</font></p>
<table border="0" width="75%" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="<% =strPopUpBorderColor %>">
    <table border="0" width="100%" cellspacing="1" cellpadding="1">
      <TR>
        <TD bgColor=<% =strPopUpTableColor %> align=center nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="aim:goIM?screenname=<% =Request.QueryString("AIM") %>" alt="����ѶϢ">����ѶϢ</a></font></td>
      </tr>
      <TR>
        <TD bgColor=<% =strPopUpTableColor %> align=center nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="aim:goChat?ROOMname=<% =Request.QueryString("AIM") %>">����������</a></font></td>
      </tr>
      <TR>
        <TD bgColor=<% =strPopUpTableColor %> align=center nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="aim:addBuddy?screenname=<% =Request.QueryString("AIM") %>">�����������</a></font></td>
      </tr>
    </TABLE>
    </td>
  </tr>
</table>

<% end select %>
<% If InStr(Request.ServerVariables("HTTP_REFERER"), "pop_profile.asp") <> 0 then %>
<p><font size="<% =strDefaultFontSize %>"><a href="JavaScript:history.go(-1)">�ص� <% =Request.QueryString("M_NAME")%> �Ļ�Ա����</a></font></p>
<% end if %>

<!--#INCLUDE FILE="inc_footer_short.asp" -->
