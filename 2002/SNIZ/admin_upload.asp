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
<!--#include file="config.asp"-->
<!--#include file="inc_functions.asp"-->
<!--#include file="inc_top_short.asp"-->
<% mLev = cint(ChkUser2(strDBNTUserName, Request.Cookies(strUniqueID & "User")("Pword"))) %>
<br><br><br>
<div align="center"> <FORM enctype="multipart/form-data" action="admin_UploadEngine.asp" method=POST name="fileup">
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
   <TABLE width="80%" border=0 cellspacing=1 cellpadding=4 align="center">
    <TR align="center" bgColor="<% =strPopUpTableColor %>">

     <TD align="center">
	 ��ѡ��Ҫ�ϴ����ļ�
      <INPUT name="Folder" type="hidden" value="<%= strDBNTUsername %>" >
      <input name="REPLY_ID" type="hidden" value="<% =Request("REPLY_ID") %>">
	  <input name="TOPIC_ID" type="hidden" value="<% =Request("TOPIC_ID") %>">
	  <input name="MEMBER_ID" type="hidden" value="<% =Request("MEMBER_ID") %>">
    </TD>
    </TR>
    <TR>

     <TD align="center">
      <INPUT name="file1" type="FILE" size=20 value="">
     </TD>
	 <td></td>
    </TR>
	<TR>
     <TD align="center"><br>
      <INPUT name="Upload" type=SUBMIT value="��ʼ�ϴ�" >
     </TD>
	 <td></td>
    </TR>
	<% If mlev = 4 then %>
		<TR>
     <TD align="center"><br>
      ���Ŀ¼ <INPUT name="uploaddir" type="text" value="tools" >
     </TD>
	 <td></td>
    </TR>
<% End If %>
  </FORM>
   </TABLE>
   </font>
<!--#include file="inc_footer_short.asp"-->

