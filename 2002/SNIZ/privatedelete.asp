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
<!--#INCLUDE FILE="inc_top.asp" -->
<%
	 
	strSql = "SELECT "   & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME,  " & strTablePrefix & "PM.M_ID,  " & strTablePrefix & "PM.M_TO "
	strSql = strSql & " FROM " & strTablePrefix & "MEMBERS , " & strTablePrefix & "PM "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & Request.Cookies(strUniqueID & "User")("Name") & "'"
	strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_PASSWORD = '" & Request.Cookies(strUniqueID & "User")("PWord") & "'"
	strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "PM.M_TO "
	strSql = strSql & " AND " & strMemberTablePrefix & "PM.M_ID =  " & Request.QueryString("id")
	
	Set rsMessage = my_Conn.Execute(strSql)
   
   
	if rsMessage.EOF or rsMessage.BOF then
			
%>
<center>
<br>
<br>
<P><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">��ûȨ�޲���ɾ���������Ļ�.</p>
<p><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript: onClick= history.go(-1)">������һҳ</a></font></p>
<br>
<br>
<% else 
	'## Forum_SQL - Delete Message
	strSql = "DELETE FROM " & strTablePrefix & "PM "
	strSql = strSql & " WHERE " & strTablePrefix & "PM.M_ID = " & Request.QueryString("id")
			
	my_Conn.Execute strSql		
%>
<meta http-equiv="Refresh" content="0; URL=pm_view.asp">
<center>
<br>
<br>
<P><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">���Ļ��Ѿ�ɾ��!</font></p>
<P><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b><a href="pm_view.asp">�������Ļ��ռ���.</a></b></font></p>
<br>
<br>
</center>
<% end if %>
</center>
<!--#INCLUDE FILE="inc_footer.asp" -->