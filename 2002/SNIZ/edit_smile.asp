<%
'########## Snitz Forums 2000 Version 3.1 SR4 ####################
'#                                                               #
'#  汉化修改: 资源搜罗站                                         #
'#  电子邮件: cgier@21cn.com                                     #
'#  主页地址: http://www.sdsea.com                               #
'#            http://www.99ss.net                                #
'#            http://www.cdown.net                               #
'#	     http://www.wzdown.com                               #
'#	     http://www.13888.net                                #
'#  论坛地址:http://ubb.yesky.net                                #
'#  最后修改日期: 2001/03/12    中文版本：Version 3.1 SR4        #
'#################################################################
'# 原始来源                                                      #
'# Snitz Forums 2000 Version 3.1 SR4                             #
'# Copyright 2000 http://forum.snitz.com - All Rights Reserved   #
'#################################################################
'#【版权声明】                                                   #
'#                                                               #
'# 本软体为共享软体(shareware)提供个人网站免费使用，请勿非法修改,#
'# 转载，散播，或用于其他图利行为，并请勿删除版权声明。          #
'# 如果您的网站正式起用了这个脚本，请您通知我们，以便我们能够知晓#
'# 如果可能，请在您的网站做上我们的链接，希望能给予合作。谢谢！  #
'#################################################################
'# 请您尊重我们的劳动和版权，不要删除以上的版权声明部分，谢谢合作#
'# 如有任何问题请到我们的论坛告诉我们                            #
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
response.write("只有论坛管理员才能修改表情设置, 请以管理员身份登陆")
%>
<!--#include file="inc_footer.asp" -->
<%end if%>
