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
Select case request.querystring("type")
	Case "cat"
    strsql = "Insert into "&strTablePrefix&"smile_cat (cat_name) Values ('"&Replace(Request.form("cat_name"), "'", "''")&"')"
    my_conn.execute(strsql)
    response.write("<center>����������ɣ�<a href=""admin_smiles.asp"">����</a>��������</center>")
    Case else
if request.form("smile_code") = "" or request.form("smile_url") = "" then
	response.write("���δ��д����ĳ�ִ���")
    my_conn.close
    set my_conn = nothing
    response.end
end if

strsql = "INSERT INTO "&strTablePrefix&"smiles (smile_code, smile_url, smile_name, cat_id) Values("
strsql = strsql & "'"&request.form("smile_code")&"' ,'"&request.form("smile_url")&"', '"&request.form("smile_name")&"', "&request.form("cat_id")&")"
my_conn.execute(strsql)

response.redirect("admin_smiles.asp")
End select
%>
<!--#INCLUDE file="inc_footer.asp" -->
<%else
response.redirect("admin_login.asp")
end if%>
