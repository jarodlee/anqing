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
    <td width="33%" align="left" nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    <img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="default.asp">������̳��ҳ</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" height=15 width=15 border="0">&nbsp;��̳��������<br>
    </font></td>
  </tr>
</table>

<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td bgcolor="<% =strTableBorderColor %>">
    <table border="0" width="100%" cellspacing="1" cellpadding="3">
      <tr>
        <td bgcolor="<% =strCategoryCellColor %>" colspan=2><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" size"<% =strDefaultFontSize %>"><b>��������</b></font></td>
      </tr>
      <tr>
        <td bgcolor="<% =strForumCellColor %>" valign=top><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
        <p><b>��̳�����趨��</b>
        <UL>
        <LI><a href="admin_config_system.asp">��������</a></LI>
        <LI><a href="admin_config_features.asp">��̳����</a></LI>
<%	If strAuthType = "nt" then %>
        <LI><a href="admin_config_NT_features.asp">NT�����趨</a></LI>
<%	End if %>
        <LI><a href="admin_config_members.asp">��Աע��ѡ������</a></LI>
        <LI><a href="admin_config_ranks.asp">�ȼ��趨</a></LI>
        <LI><a href="admin_config_datetime.asp">����������/ʱ���趨</a></LI>
        <LI><a href="admin_config_email.asp">�ʼ���������</a></LI>
        <LI><a href="admin_config_colors.asp">�������</a></LI>
        <LI><a href="admin_config_badwords.asp">�����������</a></LI>
        </UL></p>
        </font></td>
        <td bgcolor="<% =strForumCellColor %>" valign=top><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
        <p><b>�����趨ѡ���빦�ܣ�</b>
        <UL>
        <LI><a href="admin_delete_topics.asp">��Ա���ӹ�������</a></LI>
        <LI><a href="admin_smiles.asp">��������</a></LI>
        <LI><a href="admin_moderators.asp">��������</a></LI>
        <LI><a href="admin_faq.asp">�����������</a></LI>
        <LI><a href="admin_emaillist.asp">�ʼ��б�����</a></LI>
        <LI><a href="admin_info.asp">��������������</a></LI>
        <LI><a href="admin_variable_info.asp">��̳ϵͳ��Ѷ</a></LI>
        <LI><a href="admin_count.asp">������̳״̬ͳ��</a></LI>
        <LI><a href="setup.asp">��鰲װ</a><font size="-2"><b>��ÿ�θ��������ִ�б����ܣ���</b></font></LI>
        </UL></p>
        </font></td>
      </tr>
    </table>
    </td>
  </tr>
</table>
<!--#INCLUDE FILE="mods/modCMD.asp" -->
<!--#INCLUDE file="inc_footer.asp" -->
<% Else %>
<% Response.Redirect "admin_login.asp" %>
<% End IF %>
