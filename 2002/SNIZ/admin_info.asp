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
<!--#INCLUDE FILE="inc_functions.asp" -->
<!--#INCLUDE file="inc_top.asp" -->
<table border="0" width="100%">
  <tr>
	  <td bgcolor="<% =strPageBGColor %>">
	<table border="0" align=center width="100%">
	<tr>
	    <td width="33%" align="left" nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
		<img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="default.asp">������̳��ҳ
</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="admin_home.asp">��̳��������</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" height=15 width=15 border="0">&nbsp;��������������<br>
		</font></td>
	</tr>

	</table>
	</td>
  </tr>
  <tr>
    <td bgcolor="<% =strPageBGColor %>">

    <table border="0" cellspacing="1" cellpadding="1" align=center width="100%">
      <tr>
        <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>Variable Name</b></font></td>
        <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>Value</b></font></td>
      </tr>
<% for each key in Request.ServerVariables %>
      <tr>
        <td bgColor="<% =strPopUpTableColor %>" valign="top"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b><% =key %></b></font></td>
        <td bgColor="<% =strPopUpTableColor %>"><font face="courier" size="<% =strDefaultFontSize %>"><%if Request.ServerVariables(key) = "" then
Response.Write "&nbsp;"
else
Response.Write Request.Servervariables(key)
end if 
%>
        </font></td>
      </tr>
<% next %>
    </table>
    </td>
    </tr>
</table>
</p>
<!--#INCLUDE file="inc_footer.asp" -->
<% Else %>
<% Response.Redirect "admin_login.asp" %>
<% End IF %>
