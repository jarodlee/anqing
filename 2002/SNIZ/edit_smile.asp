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
if request.form("smile_id") = "" then
	response.write("Your here by mistake")
    my_conn.close
    set my_conn=nothing
    response.end
end if

Dim smile_id, smile_code, smile_url, smile_name
smile_id = request.form("smile_id")
smile_code = request.form("smile_code")
smile_url = request.form("smile_url")
smile_name = request.form("smile_name")
smile_cat = request.form("cat_id")

strsql = "UPDATE "&strTablePrefix&"smiles SET smile_code='"&smile_code&"', smile_url='"&smile_url&"', smile_name='"&smile_name&"', cat_id="&smile_cat&" WHERE id="&smile_id
my_conn.execute(strsql)

response.redirect("admin_smiles.asp")
%>
<!--#INCLUDE file="inc_footer.asp" -->
<%else
response.write("ֻ����̳����Ա�����޸ı�������, ���Թ���Ա��ݵ�½")
%>
<!--#include file="inc_footer.asp" -->
<%end if%>
