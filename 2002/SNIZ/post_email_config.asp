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
<%
	strEmailMessage = ""
	str_m_type = "newreply"
	strEmailSubject = request.form("tsubject")
	strEmailMessage = request.form("tmessage")
	if strEmailMessage <> "" then
		strsql = "Update " & strTablePrefix & "email_config Set "
	    strsql = strsql & "m_subject='" & strEmailSubject & "', "
	    strsql = strsql & "m_message='" & strEmailMessage & "' "
	    strsql = strsql & "Where m_type='" & str_m_type & "'"
	    my_conn.execute(strsql)
	end if
	strEmailMessage = ""
	str_m_type = "replynot"
	strEmailSubject = request.form("rsubject")
	strEmailMessage = request.form("rmessage")
	if strEmailMessage <> "" then
		strsql = "Update " & strTablePrefix & "email_config Set "
	    strsql = strsql & "m_subject='" & strEmailSubject & "', "
	    strsql = strsql & "m_message='" & strEmailMessage & "' "
	    strsql = strsql & "Where m_type='" & str_m_type & "'"
	    my_conn.execute(strsql)
	end if

    
    response.write("<p><center>�ʼ�֪ͨ��Ϣ�Ѿ��޸����, <a href=""admin_email_messages.asp"">������ﷵ��</a></center><p>")


%>
<!--#INCLUDE file="inc_footer.asp" -->
<% Else %>
<% Response.Redirect "admin_login.asp" %>
<% End IF %>

