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
<%
strMessagePreview = Request.Cookies("strMessagePreview")
Response.Cookies("strMessagePreview") = ""
if strMessagePreview = "" or IsNull(strMessagePreview) then
	strMessagePreview = "[center][b]<û�����ݿ���Ԥ��>[/b][/center]"
end if
%>
<!--#INCLUDE FILE="inc_top_short.asp"-->
<script language="JavaScript">
<!--
function jumpTo(s) {if (s.selectedIndex != 0) top.location.href = s.options[s.selectedIndex].value;return 1;}
// -->
</script>
<table border="0" width="95%" height="80%" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td bgcolor="<% =strTableBorderColor %>">
    <table border="0" width="100%" height="100%" cellspacing="1" cellpadding="4">
      <tr>
        <td align="center" bgcolor="<% =strHeadCellColor %>" width="<% =strTopicWidthRight %>" height="20" <% if lcase(strTopicNoWrapRight) = "1" then Response.Write(" nowrap") %>><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">Ԥ��</font></b></td>
	  </tr>
      <tr>
        <td bgcolor="<% =strForumCellColor %>" valign="top">
        <font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% =formatStr(chkString(strMessagePreview,"preview")) %></font>
        </td>
      </tr>
    </table>
    </td>
  </tr>
</table>
</div>
<!--#INCLUDE FILE="inc_footer_short.asp" -->