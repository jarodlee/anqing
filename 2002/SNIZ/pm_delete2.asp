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
<% If Request.Form("RemoveTopic") = "1" then
'Open a connection to the database
		
	'We want to delete our products.  The list of ProductIDs that need
	'to be deleted are in a comma-delimited list...
	Dim strRemoveList
	strRemoveList = Request("Remove")
	
	
	if strRemoveList = "" then 'Do Nothing
		
	Else
		

		'Now, use the SQL set notation to Remove all of the records
		'specified by strRemoveList
		Dim strSqL
		strSqL = "UPDATE " & strTablePrefix & "PM "
		strSql = strSql & "SET FORUM_PM.M_OUTBOX = 0 " & _
				 "WHERE M_ID IN (" & strRemoveList & ")"
				 
		my_Conn.Execute (strSql)
	
				
%>
<center>
<br>
<br>
<P><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><% Response.Write Request("Remove").Count & " message(s) were removed!" %></font></p>
<P><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b><a href="pm_view.asp">�������Ļ��ռ���.</a></b></font></p>
<br>
<br>
</center>
<%	
		
		
	End If


else

		
	Dim strDeleteList
	strDeleteList = Request("Delete")
	
	
	if strDeleteList = "" then
		'No messages to delete
%>
<center>
<br>
<br>
<P><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">��û��ѡ��Ҫɾ�������Ļ�!</font></p>
<P><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b><a href="pm_view.asp">�������Ļ��ռ���.</a></b></font></p>
<br>
<br>
</center>
<%
	Else
		

		'Now, use the SQL set notation to delete all of the records
		'specified by strDeleteList
		
		strSQL = "DELETE FROM FORUM_PM " & _
				 "WHERE M_ID IN (" & strDeleteList & ")"
				 
		my_Conn.Execute (strSql)
	
		
		
%>
<center>
<br>
<br>
<P><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><% Response.Write Request("Delete").Count & " Message(s) were deleted!" %></font></p>
<P><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b><a href="pm_view.asp">�������Ļ��ռ���.</a></b></font></p>
<br>
<br>
</center>
<%	

		
	End If
end if

%> 

<!--#INCLUDE FILE="inc_footer.asp" -->