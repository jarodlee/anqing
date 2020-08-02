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
<P><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b><a href="pm_view.asp">返回悄悄话收件箱.</a></b></font></p>
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
<P><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">你没有选择要删除的悄悄话!</font></p>
<P><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b><a href="pm_view.asp">返回悄悄话收件箱.</a></b></font></p>
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
<P><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b><a href="pm_view.asp">返回悄悄话收件箱.</a></b></font></p>
<br>
<br>
</center>
<%	

		
	End If
end if

%> 

<!--#INCLUDE FILE="inc_footer.asp" -->