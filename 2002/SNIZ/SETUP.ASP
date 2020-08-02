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

'#################################################################################
    strNewVersion = "Snitz Forums 2000 Version 3.1 SR4"
'#################################################################################

ResponseCode = Request.QueryString("RC")
if ResponseCode <> "" then 'No parameter
	blnSetup = "Y"
else
	strCookieURL = Left(Request.ServerVariables("Path_Info"), InstrRev(Request.ServerVariables("Path_Info"), "/"))
	Application(strCookieURL & "ConfigLoaded")= ""
end if
%>
<!--#INCLUDE FILE="config.asp" -->
<%
Response.Buffer = True
%>
<html>
<head>
<title>Forum-Setup Page</title>
<meta name="copyright" content="Snitz Forums 2000 Version 3.1 SR4 汉化修改：资源搜罗站">
<Style><!--
a:link    {color:<% =strLinkColor %>;text-decoration:<% =strLinkTextDecoration %>}
a:visited {color:<% =strVisitedLinkColor %>;text-decoration:<% =strVisitedTextDecoration %>}
a:hover   {color:<% =strHoverFontColor %>;text-decoration:<% =strHoverTextDecoration %>}
--></style>
</head>

<body bgColor="white" text="midnightblue" link="darkblue" aLink=red vLink="red" onLoad="window.focus()">
<%
if ResponseCode = "" then 'No parameter

'	Check to see if all the fields are in the database
'	FORUM_CONFIG doesn't need to be checked, that is done in CONFIG.ASP
	
	on error resume next
	
	set my_Conn = Server.CreateObject("ADODB.Connection")
	my_Conn.Open strConnString	
	
	for counter = 0 to my_conn.Errors.Count -1
		ConnErrorNumber = my_conn.Errors(counter).Number
		if ConnErrorNumber <> 0 then 
			my_conn.Errors.Clear 
			Response.Redirect "Setup.asp?RC=1&EC=" & ConnErrorNumber
		end if
	next
	
	my_Conn.Errors.Clear

	strSql = "SELECT CAT_ID, FORUM_ID, F_STATUS, F_MAIL, F_SUBJECT, F_URL, F_DESCRIPTION, F_TOPICS, F_COUNT, "
	strSql = strSql & "F_LAST_POST, F_PASSWORD_NEW, F_PRIVATEFORUMS, F_TYPE, F_IP, F_LAST_POST_AUTHOR "
	strSql = strSql &  " FROM " & strTablePrefix & "FORUM"
	
	my_Conn.Execute strSql
	
	Call CheckSqlError()
	
	my_Conn.Errors.Clear
	
	strSql = "SELECT CAT_ID, FORUM_ID, TOPIC_ID, T_STATUS, T_MAIL, T_SUBJECT, T_MESSAGE, T_AUTHOR, T_REPLIES, "
	strSql = strSql & "T_VIEW_COUNT, T_LAST_POST, T_DATE, T_LAST_POSTER, T_IP, T_LAST_POST_AUTHOR "
	strSql = strSql & " FROM " & strTablePrefix & "TOPICS"

	my_Conn.Execute strSql

	Call CheckSqlError()
	
	my_Conn.Errors.Clear
	
	strSql = "SELECT MEMBER_ID, M_STATUS, M_NAME, M_USERNAME, M_PASSWORD, M_EMAIL, M_COUNTRY, M_HOMEPAGE, "
	strSql = strSql & "M_SIG, M_DEFAULT_VIEW, M_LEVEL, M_AIM, M_YAHOO, M_ICQ, M_POSTS, M_DATE, "
	strSql = strSql & "M_LASTHEREDATE, M_LASTPOSTDATE, M_TITLE, M_SUBSCRIPTION, M_HIDE_EMAIL, "
	strSql = strSql & "M_RECEIVE_EMAIL, M_LAST_IP, M_IP, M_FIRSTNAME, M_LASTNAME, M_OCCUPATION, "
	strSql = strSql & "M_SEX, M_AGE, M_HOBBIES, M_LNEWS, M_QUOTE, M_BIO, M_MARSTATUS, "
	strSql = strSql & "M_LINK1, M_LINK2, M_CITY, M_STATE, M_PHOTO_URL FROM " & strMemberTablePrefix & "MEMBERS"
	strSql = strSql & " WHERE MEMBER_ID = 1;"

	my_Conn.Execute strSql

	Call CheckSqlError()
	
	my_Conn.Errors.Clear

	on error goto 0

	if strVersion <> strNewVersion then
		Response.Redirect "setup.asp?RC=2&MAIL=" & Server.UrlEncode(strSender) & "&VER=" & Server.URLEncode(strVersion) & "&EC=" & Server.UrlEncode("Different or New Version-ID detected")
	end if

'## This part of the code is only reached if all is ok !!
%>
<div align=center><center>
	<p><font face="宋体, Arial, Helvetica" size="4">论坛设置完毕！</font></p>
</center></div>

<table border="0" cellspacing="0" cellpadding="5" width="50%" height="50%" align="center">
	<tr>
		<td bgColor=navyblue align="center">
		<p>
		<font face="宋体, Arial, Helvetica" size="2">
		<b>恭喜!!</b><br>
		论坛已经安装成功.<br>
		现在你可以开始使用 Snitz Forums 2000 论坛了.</font></p>
		<font face="宋体, Arial, Helvetica" size="2">
		<b>记得修改默认的论坛管理密码 !!</b></font></p>
		<font face="宋体, Arial, Helvetica" size="2">
		如果你还有什么问题可以到下面的网址跟我们联系: <a href="http://forum.snitz.com">http://forum.snitz.com</a></font></p></td>
	</tr>
	<tr>
		<td align="center">
		<font face="宋体, Arial, Helvetica" size="2">
		<a href="default.asp" target="_top">点击这里进入论坛.</a>
		</font></td>
	</tr>
	<tr>
		<td align="center">
		<font face="宋体, Arial, Helvetica" size="2">
		<a href="setup.asp?RC=3" target="_top">更新数据库.</a>
		</font></td>
	</tr>
<% if strDBType <> "access" then %>
	<tr>
		<td align="center">
		<font face="宋体, Arial, Helvetica" size="2">
		<a href="setup.asp?RC=5" target="_top">建立数据库.</a>
		</font></td>
	</tr>
<% end if %>
</table>	

<%
elseif ResponseCode = 1 then '## cannot open database

ErrorCode = Request.QueryString("EC")	
CustomCode = Request.QueryString("CC")
%>
<div align=center><center>
	<p><font face="宋体, Arial, Helvetica" size="4">安装不能继续进行，发现一个错误!!</font></p>
</center></div>

<table border="0" cellspacing="0" cellpadding="5" width="50%" height="50%" align="center">
	<tr>
		<td bgColor=navyblue align="center">
		<p>
		<font face="宋体, Arial, Helvetica" size="2">
<%
if CustomCode = 1 then
%>
		错误：不能打开数据库!!<br>
		检查你的 config.asp 文件设置好<br><b>strConnString</b>连接数据库参数.<br>
		<br>
<%
elseif CustomCode = 2 then
%>
		不能从数据库中读取 tables.<br> Make sure none of the tables are exclusively locked by another user.<br><br>
<%
elseif CustomCode = 3 then 
%>
		错误：不能打开数据库!!<br> 请确定你的数据库用户名跟连接密码是正确的.<br><br>
<%
else
%>
		错误：不能打开数据库!!<br>
		<br>
<%

end if
		if ErrorCode <> "" then 
			Response.Write("<p>Code :  " & ErrorCode & "</p>")
		end if
%>
		</font></p></td>
	</tr>
	<tr>
		<td align="center">
		<font face="宋体, Arial, Helvetica" size="2">
<% 
if Request.QueryString("RET") <> "" then 
%>
		<a href="<%=Request.QueryString("RET")%>" target="_top">点击这里返回上一页.</a>
<%
else
%>
		<a href="setup.asp" target="_top">点击这里重试.</a>
<%
end if
%> 
		</font></td>
	</tr>
</table>	
<%
elseif ResponseCode = 2 then '## cannot find all the fields in the database

strSender = Request.QueryString("MAIL")
strVersion = Request.QueryString("VER")
ErrorCode = Request.QueryString("EC")
CustomCode = Request.QueryString("CC")

if ErrorCode = "-2147467259" then
	Response.Redirect "setup.asp?RC=5&EC=" & ErrorCode
	Response.End
elseif ErrorCode = "-2147217865" then
	if strDBType = "access" then
		Response.Redirect "setup.asp?RC=1&EC=" & server.URLEncode("The Access database has missing tables !!")
		Response.End
	else
		Session(strCookieURL & "Approval") = "15916941253"
%>
<div align=center><center>
	<p><font face="宋体, Arial, Helvetica" size="4">The database needs to be installed !!</font></p>
</center></div>
<table border="0" cellspacing="0" cellpadding="5" width="50%" height="50%" align="center">
	<tr>
		<td bgColor=navyblue align="center">
		<p>
		<font face="宋体, Arial, Helvetica" size="2">
		使用论坛前请先建立数据库.<br>
		</font></p></td>
	</tr>
	<tr>
		<td align="center">
		<font face="宋体, Arial, Helvetica" size="2">
		<a href="setup.asp?RC=5&strDBType=<% =strDBType %>"  target="_top">点击这里建立一个数据库.</a><br><br>
		<a href="setup.asp" target="_top">点击这里重试.</a>
		</font></td>
	</tr>
</table>
<%
	end if
else
%>
<div align=center><center>
	<p><font face="宋体, Arial, Helvetica" size="4">The database needs to be upgraded !!</font></p>
</center></div>
<table border="0" cellspacing="0" cellpadding="5" width="50%" height="50%" align="center">
	<tr>
		<td bgColor=navyblue align="center">
		<p>
		<font face="宋体, Arial, Helvetica" size="2">
		The database you are using needs to be upgraded !!<br>
<%
	if MAIL <> ""then 
%>
		If you are not a Administrator at this forum<br> please report this error here: <a href="mailto:<% =strSender %>"><% =strSender %></a>.<br><br>
<%
	end if 
	if ErrorCode <> "" then 
			Response.Write("<p>Code :  " & ErrorCode & "</p>" )
	end if
%>
		</font></p></td>
	</tr>
	<tr>
		<td align="center">
		<font face="宋体, Arial, Helvetica" size="2">
		<a href="setup.asp?RC=3&MAIL=<% =Server.URLEncode(strSender) %>"  target="_top">Click here to upgrade the database.</a><br><br>
		<a href="default.asp" target="_top">Click here to retry.</a>
		</font></td>
	</tr>
</table>	
<%
end if


elseif ResponseCode = 3 then '## upgrade database

	if Session(strCookieURL & "Approval") = "15916941253" Then

	'## logon was ok proceed with upgrade
%>	
	<div align=center><center>
<% if strDBType = "sqlserver" then %>
	<p><font face="宋体, Arial, Helvetica" size="4">Select the SQL-Server upgrade options.</font></p>
<% elseif strDBType = "mysql" then  %>
	<p><font face="宋体, Arial, Helvetica" size="4">MySql database upgrade.</font></p>
<% else %>
	<p><font face="宋体, Arial, Helvetica" size="4">Access 97 database upgrade</font></p>
<% end if %>
</center></div>

<table border="0" cellspacing="0" cellpadding="5" width="50%" height="50%" align="center">
	<tr>
		<td bgColor=navyblue <% if strDBType = "sqlserver" then Response.Write("align=""left""") else Response.Write("align=""center""") end if %>>
		<p>
		<font face="宋体, Arial, Helvetica" size="2">
		Select the version you want to upgrade from:
		</p>
        <form action="setup.asp?RC=4&strDBType=<% =strDBType %>" method="post" id=Form1 name=Form1>
          <p><select size="1" name="OldVersion">
            <option value="5" selected>Snitz Forums 2000 V3.1 Service Release 3 Final</option>
            <option value="5">Snitz Forums 2000 V3.1 Service Release 3 Beta 4</option>
            <option value="4">Snitz Forums 2000 V3.1 Service Release 3 Beta 2</option>
            <option value="3">Snitz Forums 2000 V3.1 Service Release 3 Beta 1</option>
            <option value="3">Snitz Forums 2000 V3.1 Service Release 2</option>
            <option value="3">Snitz Forums 2000 V3.1 Service Release 1</option>
            <option value="3">Snitz Forums 2000 V3.1 Final</option>
            <option value="3">Snitz Forums 2000 V3.1 Beta 5</option>
            <option value="3">Snitz Forums 2000 V3.1 Beta 4</option>
			<option value="2">Snitz Forums 2000 V3.1 Beta 3</option>
			<option value="1">Snitz Forums 2000 V3.1 Beta 2</option>
            <option value="0">Snitz Forums 2000 V3.0 Service Release 2</option>
            <option value="0">Snitz Forums 2000 V3.0 Service Release 1</option>
            <option value="0">Snitz Forums 2000 Version 3.0 Final</option>
            <option value="0">Snitz Forums 2000 V3 RC5</option>
            <option value="0">Snitz Forums 2000 V3 RC4</option>
            <option value="0">Snitz Forums 2000 V3 RC3</option>
            <option value="0">Snitz Forums 2000 V3 RC2</option>
            <option value="0">Snitz Forums 2000 V3 RC1</option>            
          </select></p>
	  <p>
	  <font face="宋体, Arial, Helvetica" size="2">
	  Current database versionstring =<br><b> <% =strVersion %></b>
	  </p>	
          <% if strDBType = "sqlserver" then %>
          <p>Select the SQL-server version you are using:</p>
          <p><input type="radio" class="radio" name="SQL_Server" value="SQL6">SQL-Server 6.5<br>
          <input type="radio" class="radio" checked name="SQL_Server" value="SQL7">SQL-Server 7</p>
          <% end if %>
          <% if strDBType <> "access" then %>
          <p>To upgrade the database you need to provide a username
          and password of a user that has table creation/modification rights at the database you use.<br>
          This might not be the same user as you use in your connectionstring !<br>
          <br>
          Name:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input type="text" name="DBUserName" size="20"><br>
          Password: <input type="password" name="DBPassword" size="20"></p>
          <% end if %>
          </font>
        </td>
		</tr>	
		<tr>
			  <td align="center"><input type="submit" value="Continue" id="Submit1" name="Submit1"></td>
		</tr>
		</form>
</table>	
	
<%	
	else
	strSender = Request.QueryString("MAIL")
%>
<div align=center><center>
	<p><font face="宋体, Arial, Helvetica" size="4">You need to logon first.</font></p>
</center></div>

<form action="setup_login.asp" method="post" id=Form1 name=Form1>
<input type="hidden" name="setup" value="Y">
	<table border="0" cellspacing="0" cellpadding="5" width="50%" height="50%" align="center">
		<tr>
			<td bgColor=navyblue align="left">
			<p>
			<font face="宋体, Arial, Helvetica" size="2">
			To upgrade the database you need to be logged on as a forum administrator.<br>
<%
if strSender <> "" then
%>
			If you are not the Administrator of this forum<br> please report this error here: <a href="mailto:<% =strSender %>"><% =strSender %></a>.<br><br>
<%
end if
%>
			</font></p></td>
		</tr>
		<tr>
			<td>
			<table border="0" cellspacing="0" cellpadding="5" border="0" cellspacing="2" cellpadding="0" align="center">
			<tr>
			  <td align="center" colspan="2" bgColor=navyblue><b><font face="宋体, Arial, Helvetica" size="2">管理者登入</font></b></td>
			</tr>
			<tr>
			  <td align="right" nowrap><b><font face="宋体, Arial, Helvetica" size="2">使用者名称：</font></b></td>
			  <td><input type="text" name="Name"></td>
			</tr>
			<tr>
			  <td align="right" nowrap nowrap><b><b><font face="宋体, Arial, Helvetica" size="2">密码：</font></b></td>
			  <td><input type="Password" name="Password"></td>
			</tr>
			<tr>
			  <td colspan="2" align="right"><input type="submit" value="登入" id="Submit1" name="Submit1"></td>
			</tr>
			</table>
			</td>
		</tr>
		</table>
</form>

</font>
<%
	end if
	
elseif ResponseCode = 4 then '## start upgrading database

	if Session(strCookieURL & "Approval") = "15916941253" Then

	'## logon was ok proceed with upgrade
%>	
	<div align=center><center>
	<p><font face="宋体, Arial, Helvetica" size="4">Please Wait until the upgrade has been completed !</font></p>
<%
	on error resume next
	
	strUpgradeString = strConnString
	strSQL_Server = Request.Form("Sql_Server")
	
	if strDBType <> "access" then
		strUpgradeString = CreateConnectionString(strConnString, Request.Form("DBUserName"), Request.Form("DBPassword"))
	end if
	
	set my_Conn = Server.CreateObject("ADODB.Connection")
	my_Conn.Open strUpgradeString
	
	for counter = 0 to my_conn.Errors.Count -1
		ConnErrorNumber = my_conn.Errors(counter).Number
		if ConnErrorNumber <> 0 then 
			my_conn.Errors.Clear 
			Response.Redirect "setup.asp?RC=1&EC=" & ConnErrorNumber & "&RET=" & Server.URLEncode("setup.asp?RC=3")
		end if
	next
	
	on error goto 0
	
	dim intCriticalErrors, intWarnings, Prefix, FieldName, TableName, DataType
	intCriticalErrors = 0
	intWarnings = 0
	
	Prefix	  = 1
	FieldName = 2
	TableName = 3
	DataType_Access  = 4
	DataType_SQL6 = 5
	DataType_SQL7 = 6
	DataType_MySQL = 7 
	ConstraintAccess = 8
	ConstraintSQL6 = 9 
	ConstraintSQL7 = 10
	ConstraintMySQL = 11
	Access = 1
	SQL6 = 2
	SQL7 = 3
	MySql = 4
			
	if not(IsNull(Request.Form("OldVersion"))) then
		OldVersion = Request.Form("OldVersion")
	else 
		OldVersion = Request.QueryString("OldVersion")
	end if
	
	if OldVersion = 0 then

		Dim NewColumns(8,11)
	
		NewColumns(0, Prefix)	 = strTablePrefix
		NewColumns(0, FieldName) = "C_STRSHOWSTATISTICS"
		NewColumns(0, TableName) = "CONFIG"
		NewColumns(0, DataType_Access) = "SMALLINT"
		NewColumns(0, DataType_SQL6) = "SMALLINT"
		NewColumns(0, DataType_SQL7) = "SMALLINT"
		NewColumns(0, DataType_MySQL) = "SMALLINT"
		NewColumns(0, ConstraintAccess)  = "NULL"
		NewColumns(0, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0362 DEFAULT 1"
		NewColumns(0, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0362 DEFAULT 1"
		NewColumns(0, ConstraintMySQL)  = "DEFAULT 1 NULL"
					
		NewColumns(1, Prefix)	 = strTablePrefix
		NewColumns(1, FieldName) = "C_STRSHOWIMAGEPOWEREDBY"
		NewColumns(1, TableName) = "CONFIG"
		NewColumns(1, DataType_Access) = "SMALLINT"
		NewColumns(1, DataType_SQL6) = "SMALLINT"
		NewColumns(1, DataType_SQL7) = "SMALLINT"
		NewColumns(1, DataType_MySQL) = "SMALLINT"
		NewColumns(1, ConstraintAccess)  = "NULL"
		NewColumns(1, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0363 DEFAULT 1"
		NewColumns(1, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0363 DEFAULT 1"
		NewColumns(1, ConstraintMySQL)  = "DEFAULT 1 NULL"
		
		NewColumns(2, Prefix)	 = strTablePrefix
		NewColumns(2, FieldName) = "C_STRLOGONFORMAIL"
		NewColumns(2, TableName) = "CONFIG"
		NewColumns(2, DataType_Access) = "SMALLINT"
		NewColumns(2, DataType_SQL6) = "SMALLINT"
		NewColumns(2, DataType_SQL7) = "SMALLINT"
		NewColumns(2, DataType_MySQL) = "SMALLINT"
		NewColumns(2, ConstraintAccess)  = "NULL"
		NewColumns(2, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0364 DEFAULT 1"
		NewColumns(2, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0364 DEFAULT 1"
		NewColumns(2, ConstraintMySQL)  = "DEFAULT 1 NULL"
		
		NewColumns(3, Prefix)	 = strTablePrefix
		NewColumns(3, FieldName) = "C_STRSHOWPAGING"
		NewColumns(3, TableName) = "CONFIG"
		NewColumns(3, DataType_Access) = "SMALLINT"
		NewColumns(3, DataType_SQL6) = "SMALLINT"
		NewColumns(3, DataType_SQL7) = "SMALLINT"
		NewColumns(3, DataType_MySQL) = "SMALLINT"
		NewColumns(3, ConstraintAccess)  = "NULL"
		NewColumns(3, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0365 DEFAULT 0"
		NewColumns(3, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0365 DEFAULT 0"
		NewColumns(3, ConstraintMySQL)  = "DEFAULT 0 NULL"
			
		NewColumns(4, Prefix)	 = strTablePrefix
		NewColumns(4, FieldName) = "C_STRSHOWTOPICNAV"
		NewColumns(4, TableName) = "CONFIG"
		NewColumns(4, DataType_Access) = "SMALLINT"
		NewColumns(4, DataType_SQL6) = "SMALLINT"
		NewColumns(4, DataType_SQL7) = "SMALLINT"
		NewColumns(4, DataType_MySQL) = "SMALLINT"
		NewColumns(4, ConstraintAccess)  = "NULL"
		NewColumns(4, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0366 DEFAULT 0"
		NewColumns(4, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0366 DEFAULT 0"
		NewColumns(4, ConstraintMySQL)  = "DEFAULT 0 NULL"
			
		NewColumns(5, Prefix)	 = strTablePrefix
		NewColumns(5, FieldName) = "C_STRPAGESIZE"
		NewColumns(5, TableName) = "CONFIG"
		NewColumns(5, DataType_Access) = "SMALLINT"
		NewColumns(5, DataType_SQL6) = "SMALLINT"
		NewColumns(5, DataType_SQL7) = "SMALLINT"
		NewColumns(5, DataType_MySQL) = "SMALLINT"
		NewColumns(5, ConstraintAccess)  = "NULL"
		NewColumns(5, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0367 DEFAULT 15"
		NewColumns(5, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0367 DEFAULT 15"
		NewColumns(5, ConstraintMySQL)  = "DEFAULT 15 NULL"
			
		NewColumns(6, Prefix)	 = strTablePrefix
		NewColumns(6, FieldName) = "C_STRPAGENUMBERSIZE"
		NewColumns(6, TableName) = "CONFIG"
		NewColumns(6, DataType_Access) = "SMALLINT"
		NewColumns(6, DataType_SQL6) = "SMALLINT"
		NewColumns(6, DataType_SQL7) = "SMALLINT"
		NewColumns(6, DataType_MySQL) = "SMALLINT"
		NewColumns(6, ConstraintAccess)  = "NULL"
		NewColumns(6, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0368 DEFAULT 10"
		NewColumns(6, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0368 DEFAULT 10"
		NewColumns(6, ConstraintMySQL)  = "DEFAULT 10 NULL"
		
		NewColumns(7, Prefix)	 = strTablePrefix
		NewColumns(7, FieldName) = "F_LAST_POST_AUTHOR"
		NewColumns(7, TableName) = "FORUM"
		NewColumns(7, DataType_Access) = "INT"
		NewColumns(7, DataType_SQL6) = "INT"
		NewColumns(7, DataType_SQL7) = "INT"
		NewColumns(7, DataType_MySQL) = "INT"
		NewColumns(7, ConstraintAccess)  = "NULL"
		NewColumns(7, ConstraintSQL6)  = "NULL"
		NewColumns(7, ConstraintSQL7)  = "NULL"
		NewColumns(7, ConstraintMySQL)  = "NULL"
		
		NewColumns(8, Prefix)	 = strTablePrefix
		NewColumns(8, FieldName) = "T_LAST_POST_AUTHOR"
		NewColumns(8, TableName) = "TOPICS"
		NewColumns(8, DataType)  = "INT"
		NewColumns(8, DataType_Access) = "INT"
		NewColumns(8, DataType_SQL6) = "INT"
		NewColumns(8, DataType_SQL7) = "INT"
		NewColumns(8, DataType_MySQL) = "INT"
		NewColumns(8, ConstraintAccess)  = "NULL"
		NewColumns(8, ConstraintSQL6)  = "NULL"
		NewColumns(8, ConstraintSQL7)  = "NULL"
		NewColumns(8, ConstraintMySQL)  = "NULL"
		
			
		call AddColumns(NewColumns, intCriticalErrors, intWarnings)
	
		'## Forum_SQL
		strSql = "UPDATE " & strTablePrefix & "CONFIG "
		strSql = strSql & " SET C_STRSHOWSTATISTICS        =  " & 1
		strSql = strSql & " ,    C_STRSHOWIMAGEPOWEREDBY    =  " & 1
		strSql = strSql & " ,    C_STRLOGONFORMAIL          =  " & 1
		strSql = strSql & " ,    C_STRSHOWPAGING            =  " & 0
		strSql = strSql & " ,    C_STRSHOWTOPICNAV          =  " & 0
		strSql = strSql & " ,    C_STRPAGESIZE              =  " & 15
		strSql = strSql & " ,    C_STRPAGENUMBERSIZE        =  " & 10
		strSql = strSql & " ,    C_STRVERSION               =  '" & strNewVersion & "'"
		strSql = strSql & " WHERE CONFIG_ID = " & 1
	
		on error resume next
		my_Conn.Errors.Clear
		my_Conn.Execute (strSql)
	
		Response.Write("<table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""50%"" align=""center"">")
	
		UpdateErrorCode = UpdateErrorCheck()
	
		on error goto 0
	
		if UpdateErrorCode = 0 then
			Response.Write("<tr><td bgColor=green align=""left"" width=""30%""><font face=""宋体, Arial, Helvetica"" size=""2""><b>Upgraded: </b></font></td><td bgColor=navyblue align=""left""><font face=""宋体, Arial, Helvetica"" size=""2""> Added default values for new fields in CONFIG table</font></td></tr>")
		elseif UpdateErrorCode = 2 then
			Response.Write("<tr><td bgColor=red align=""left"" width=""30%""><font face=""宋体, Arial, Helvetica"" size=""2""><b>Critical error: </b></font></td><td bgColor=navyblue align=""left""><font face=""宋体, Arial, Helvetica"" size=""2"">Can't add default values for new fields in CONFIG table!<b></font></td></tr>")
			intCriticalErrors = intCriticalErrors + 1
		else
			Response.Write("<tr><td bgColor=red align=""left"" width=""30%""><font face=""宋体, Arial, Helvetica"" size=""2""><b>Critical error: code: </b></font></td><td bgColor=navyblue align=""left""><font face=""宋体, Arial, Helvetica"" size=""2"">" & UpdateErrorCode & " while trying to add default values to the CONFIG table</font></td></tr>")
			intCriticalErrors = intCriticalErrors + 1
		end if
		Response.Write("</table>")
		Response.Flush
	
	end if
	if (OldVersion <= 2) then

		Dim NewColumns2(29,11)
			
		NewColumns2(0, Prefix)	 = strMemberTablePrefix
		NewColumns2(0, FieldName) = "M_FIRSTNAME"
		NewColumns2(0, TableName) = "MEMBERS"
		NewColumns2(0, DataType_Access)  = "TEXT (100)"
		NewColumns2(0, DataType_SQL6)  = "VARCHAR (100)"
		NewColumns2(0, DataType_SQL7)  = "NVARCHAR (100)"
		NewColumns2(0, DataType_MYSQL)  = "VARCHAR (100)"
		NewColumns2(0, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0369 DEFAULT ''"
		NewColumns2(0, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0369 DEFAULT ''"
		NewColumns2(0, ConstraintMySQL)  = "DEFAULT '' NULL"

		NewColumns2(1, Prefix)	 = strMemberTablePrefix
		NewColumns2(1, FieldName) = "M_LASTNAME"
		NewColumns2(1, TableName) = "MEMBERS"
		NewColumns2(1, DataType_Access)  = "TEXT (100)"
		NewColumns2(1, DataType_SQL6)  = "VARCHAR (100)"
		NewColumns2(1, DataType_SQL7)  = "NVARCHAR (100)"
		NewColumns2(1, DataType_MYSQL)  = "VARCHAR (100)"
		NewColumns2(1, ConstraintAccess)  = "NULL"
		NewColumns2(1, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0370 DEFAULT ''"
		NewColumns2(1, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0370 DEFAULT ''"
		NewColumns2(1, ConstraintMySQL)  = "DEFAULT '' NULL"

		NewColumns2(2, Prefix)	 = strMemberTablePrefix
		NewColumns2(2, FieldName) = "M_OCCUPATION"
		NewColumns2(2, TableName) = "MEMBERS"
		NewColumns2(2, DataType_Access)  = "TEXT (255)"
		NewColumns2(2, DataType_SQL6)  = "VARCHAR (255)"
		NewColumns2(2, DataType_SQL7)  = "NVARCHAR (255)"
		NewColumns2(2, DataType_MYSQL)  = "VARCHAR (255)"
		NewColumns2(2, ConstraintAccess)  = "NULL"
		NewColumns2(2, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0371 DEFAULT ''"
		NewColumns2(2, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0371 DEFAULT ''"
		NewColumns2(2, ConstraintMySQL)  = "DEFAULT '' NULL"

		NewColumns2(3, Prefix)	 = strMemberTablePrefix
		NewColumns2(3, FieldName) = "M_SEX"
		NewColumns2(3, TableName) = "MEMBERS"
		NewColumns2(3, DataType_Access)  = "TEXT (50)"
		NewColumns2(3, DataType_SQL6)  = "VARCHAR (50)"
		NewColumns2(3, DataType_SQL7)  = "NVARCHAR (50)"
		NewColumns2(3, DataType_MYSQL)  = "VARCHAR (50)"
		NewColumns2(3, ConstraintAccess)  = "NULL"
		NewColumns2(3, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0372 DEFAULT ''"
		NewColumns2(3, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0372 DEFAULT ''"
		NewColumns2(3, ConstraintMySQL)  = "DEFAULT '' NULL"

		NewColumns2(4, Prefix)	 = strMemberTablePrefix
		NewColumns2(4, FieldName) = "M_AGE"
		NewColumns2(4, TableName) = "MEMBERS"
		NewColumns2(4, DataType_Access)  = "TEXT (10)"
		NewColumns2(4, DataType_SQL6)  = "VARCHAR (10)"
		NewColumns2(4, DataType_SQL7)  = "NVARCHAR (10)"
		NewColumns2(4, DataType_MYSQL)  = "VARCHAR (10)"
		NewColumns2(4, ConstraintAccess)  = "NULL"
		NewColumns2(4, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0373 DEFAULT ''"
		NewColumns2(4, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0373 DEFAULT ''"
		NewColumns2(4, ConstraintMySQL)  = "DEFAULT '' NULL"

		NewColumns2(5, Prefix)	 = strMemberTablePrefix
		NewColumns2(5, FieldName) = "M_HOBBIES"
		NewColumns2(5, TableName) = "MEMBERS"
		NewColumns2(5, DataType_Access)  = "MEMO"
		NewColumns2(5, DataType_SQL6)  = "TEXT"
		NewColumns2(5, DataType_SQL7)  = "NTEXT"
		NewColumns2(5, DataType_MYSQL)  = "TEXT"
		NewColumns2(5, ConstraintAccess)  = "NULL"
		NewColumns2(5, ConstraintSQL6)  = "NULL"
		NewColumns2(5, ConstraintSQL7)  = "NULL"
		NewColumns2(5, ConstraintMySQL)  = "NULL"

		NewColumns2(6, Prefix)	 = strMemberTablePrefix
		NewColumns2(6, FieldName) = "M_LNEWS"
		NewColumns2(6, TableName) = "MEMBERS"
		NewColumns2(6, DataType_Access)  = "MEMO"
		NewColumns2(6, DataType_SQL6)  = "TEXT"
		NewColumns2(6, DataType_SQL7)  = "NTEXT"
		NewColumns2(6, DataType_MYSQL)  = "TEXT"
		NewColumns2(6, ConstraintAccess)  = "NULL"
		NewColumns2(6, ConstraintSQL6)  = "NULL"
		NewColumns2(6, ConstraintSQL7)  = "NULL"
		NewColumns2(6, ConstraintMySQL)  = "NULL"

		NewColumns2(7, Prefix)	 = strMemberTablePrefix
		NewColumns2(7, FieldName) = "M_QUOTE"
		NewColumns2(7, TableName) = "MEMBERS"
		NewColumns2(7, DataType_Access)  = "MEMO"
		NewColumns2(7, DataType_SQL6)  = "TEXT"
		NewColumns2(7, DataType_SQL7)  = "NTEXT"
		NewColumns2(7, DataType_MYSQL)  = "TEXT"
		NewColumns2(7, ConstraintAccess)  = "NULL"
		NewColumns2(7, ConstraintSQL6)  = "NULL"
		NewColumns2(7, ConstraintSQL7)  = "NULL"
		NewColumns2(7, ConstraintMySQL)  = "NULL"

		NewColumns2(8, Prefix)	 = strMemberTablePrefix
		NewColumns2(8, FieldName) = "M_BIO"
		NewColumns2(8, TableName) = "MEMBERS"
		NewColumns2(8, DataType_Access)  = "MEMO"
		NewColumns2(8, DataType_SQL6)  = "TEXT"
		NewColumns2(8, DataType_SQL7)  = "NTEXT"
		NewColumns2(8, DataType_MYSQL)  = "TEXT"
		NewColumns2(8, ConstraintAccess)  = "NULL"
		NewColumns2(8, ConstraintSQL6)  = "NULL"
		NewColumns2(8, ConstraintSQL7)  = "NULL"
		NewColumns2(8, ConstraintMySQL)  = "NULL"

		NewColumns2(9, Prefix)	 = strMemberTablePrefix
		NewColumns2(9, FieldName) = "M_MARSTATUS"
		NewColumns2(9, TableName) = "MEMBERS"
		NewColumns2(9, DataType_Access)  = "TEXT (100)"
		NewColumns2(9, DataType_SQL6)  = "VARCHAR (100)"
		NewColumns2(9, DataType_SQL7)  = "NVARCHAR (100)"
		NewColumns2(9, DataType_MYSQL)  = "VARCHAR (100)"
		NewColumns2(9, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0374 DEFAULT ''"
		NewColumns2(9, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0374 DEFAULT ''"
		NewColumns2(9, ConstraintMySQL)  = "DEFAULT '' NULL"

		NewColumns2(10, Prefix)	 = strMemberTablePrefix
		NewColumns2(10, FieldName) = "M_LINK1"
		NewColumns2(10, TableName) = "MEMBERS"
		NewColumns2(10, DataType_Access)  = "TEXT (255)"
		NewColumns2(10, DataType_SQL6)  = "VARCHAR (255)"
		NewColumns2(10, DataType_SQL7)  = "NVARCHAR (255)"
		NewColumns2(10, DataType_MYSQL)  = "VARCHAR (255)"
		NewColumns2(10, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0375 DEFAULT ''"
		NewColumns2(10, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0375 DEFAULT ''"
		NewColumns2(10, ConstraintMySQL)  = "DEFAULT '' NULL"

		NewColumns2(11, Prefix)	 = strMemberTablePrefix
		NewColumns2(11, FieldName) = "M_LINK2"
		NewColumns2(11, TableName) = "MEMBERS"
		NewColumns2(11, DataType_Access)  = "TEXT (255)"
		NewColumns2(11, DataType_SQL6)  = "VARCHAR (255)"
		NewColumns2(11, DataType_SQL7)  = "NVARCHAR (255)"
		NewColumns2(11, DataType_MYSQL)  = "VARCHAR (255)"
		NewColumns2(11, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0376 DEFAULT ''"
		NewColumns2(11, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0376 DEFAULT ''"
		NewColumns2(11, ConstraintMySQL)  = "DEFAULT '' NULL"

		NewColumns2(12, Prefix)	 = strMemberTablePrefix
		NewColumns2(12, FieldName) = "M_CITY"
		NewColumns2(12, TableName) = "MEMBERS"
		NewColumns2(12, DataType_Access)  = "TEXT (100)"
		NewColumns2(12, DataType_SQL6)  = "VARCHAR (100)"
		NewColumns2(12, DataType_SQL7)  = "NVARCHAR (100)"
		NewColumns2(12, DataType_MYSQL)  = "VARCHAR (100)"
		NewColumns2(12, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0377 DEFAULT ''"
		NewColumns2(12, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0377 DEFAULT ''"
		NewColumns2(12, ConstraintMySQL)  = "DEFAULT '' NULL"
			
		NewColumns2(13, Prefix)	 = strMemberTablePrefix
		NewColumns2(13, FieldName) = "M_PHOTO_URL"
		NewColumns2(13, TableName) = "MEMBERS"
		NewColumns2(13, DataType_Access)  = "TEXT (255)"
		NewColumns2(13, DataType_SQL6)  = "VARCHAR (255)"
		NewColumns2(13, DataType_SQL7)  = "NVARCHAR (255)"
		NewColumns2(13, DataType_MYSQL)  = "VARCHAR (255)"
		NewColumns2(13, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0378 DEFAULT ''"
		NewColumns2(13, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0378 DEFAULT ''"
		NewColumns2(13, ConstraintMySQL)  = "DEFAULT '' NULL"
		
		NewColumns2(14, Prefix)	 = strMemberTablePrefix
		NewColumns2(14, FieldName) = "M_STATE"
		NewColumns2(14, TableName) = "MEMBERS"
		NewColumns2(14, DataType_Access)  = "TEXT (100)"
		NewColumns2(14, DataType_SQL6)  = "VARCHAR (100)"
		NewColumns2(14, DataType_SQL7)  = "NVARCHAR (100)"
		NewColumns2(14, DataType_MYSQL)  = "VARCHAR (100)"
		NewColumns2(14, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0379 DEFAULT ''"
		NewColumns2(14, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC0379 DEFAULT ''"
		NewColumns2(14, ConstraintMySQL)  = "DEFAULT '' NULL"
		
		NewColumns2(15, Prefix)	 = strTablePrefix
		NewColumns2(15, FieldName) = "C_STRFULLNAME"
		NewColumns2(15, TableName) = "CONFIG"
		NewColumns2(15, DataType_Access) = "SMALLINT"
		NewColumns2(15, DataType_SQL6) = "SMALLINT"
		NewColumns2(15, DataType_SQL7) = "SMALLINT"
		NewColumns2(15, DataType_MySQL) = "SMALLINT"
		NewColumns2(15, ConstraintAccess)  = "NULL"
		NewColumns2(15, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1000 DEFAULT 0"
		NewColumns2(15, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1000 DEFAULT 0"
		NewColumns2(15, ConstraintMySQL)  = "DEFAULT 0 NULL"
		
		NewColumns2(16, Prefix)	 = strTablePrefix
		NewColumns2(16, FieldName) = "C_STRPICTURE"
		NewColumns2(16, TableName) = "CONFIG"
		NewColumns2(16, DataType_Access) = "SMALLINT"
		NewColumns2(16, DataType_SQL6) = "SMALLINT"
		NewColumns2(16, DataType_SQL7) = "SMALLINT"
		NewColumns2(16, DataType_MySQL) = "SMALLINT"
		NewColumns2(16, ConstraintAccess)  = "NULL"
		NewColumns2(16, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1001 DEFAULT 0"
		NewColumns2(16, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1001 DEFAULT 0"
		NewColumns2(16, ConstraintMySQL)  = "DEFAULT 0 NULL"
		
		NewColumns2(17, Prefix)	 = strTablePrefix
		NewColumns2(17, FieldName) = "C_STRSEX"
		NewColumns2(17, TableName) = "CONFIG"
		NewColumns2(17, DataType_Access) = "SMALLINT"
		NewColumns2(17, DataType_SQL6) = "SMALLINT"
		NewColumns2(17, DataType_SQL7) = "SMALLINT"
		NewColumns2(17, DataType_MySQL) = "SMALLINT"
		NewColumns2(17, ConstraintAccess)  = "NULL"
		NewColumns2(17, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1002 DEFAULT 0"
		NewColumns2(17, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1002 DEFAULT 0"
		NewColumns2(17, ConstraintMySQL)  = "DEFAULT 0 NULL"

		NewColumns2(18, Prefix)	 = strTablePrefix
		NewColumns2(18, FieldName) = "C_STRCITY"
		NewColumns2(18, TableName) = "CONFIG"
		NewColumns2(18, DataType_Access) = "SMALLINT"
		NewColumns2(18, DataType_SQL6) = "SMALLINT"
		NewColumns2(18, DataType_SQL7) = "SMALLINT"
		NewColumns2(18, DataType_MySQL) = "SMALLINT"
		NewColumns2(18, ConstraintAccess)  = "NULL"
		NewColumns2(18, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1003 DEFAULT 0"
		NewColumns2(18, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1003 DEFAULT 0"
		NewColumns2(18, ConstraintMySQL)  = "DEFAULT 0 NULL"

		NewColumns2(19, Prefix)	 = strTablePrefix
		NewColumns2(19, FieldName) = "C_STRSTATE"
		NewColumns2(19, TableName) = "CONFIG"
		NewColumns2(19, DataType_Access) = "SMALLINT"
		NewColumns2(19, DataType_SQL6) = "SMALLINT"
		NewColumns2(19, DataType_SQL7) = "SMALLINT"
		NewColumns2(19, DataType_MySQL) = "SMALLINT"
		NewColumns2(19, ConstraintAccess)  = "NULL"
		NewColumns2(19, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1004 DEFAULT 0"
		NewColumns2(19, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1004 DEFAULT 0"
		NewColumns2(19, ConstraintMySQL)  = "DEFAULT 0 NULL"

		NewColumns2(20, Prefix)	 = strTablePrefix
		NewColumns2(20, FieldName) = "C_STRAGE"
		NewColumns2(20, TableName) = "CONFIG"
		NewColumns2(20, DataType_Access) = "SMALLINT"
		NewColumns2(20, DataType_SQL6) = "SMALLINT"
		NewColumns2(20, DataType_SQL7) = "SMALLINT"
		NewColumns2(20, DataType_MySQL) = "SMALLINT"
		NewColumns2(20, ConstraintAccess)  = "NULL"
		NewColumns2(20, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1005 DEFAULT 0"
		NewColumns2(20, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1005 DEFAULT 0"
		NewColumns2(20, ConstraintMySQL)  = "DEFAULT 0 NULL"
		
		NewColumns2(21, Prefix)	 = strTablePrefix
		NewColumns2(21, FieldName) = "C_STRCOUNTRY"
		NewColumns2(21, TableName) = "CONFIG"
		NewColumns2(21, DataType_Access) = "SMALLINT"
		NewColumns2(21, DataType_SQL6) = "SMALLINT"
		NewColumns2(21, DataType_SQL7) = "SMALLINT"
		NewColumns2(21, DataType_MySQL) = "SMALLINT"
		NewColumns2(21, ConstraintAccess)  = "NULL"
		NewColumns2(21, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1006 DEFAULT 0"
		NewColumns2(21, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1006 DEFAULT 0"
		NewColumns2(21, ConstraintMySQL)  = "DEFAULT 0 NULL"

		NewColumns2(22, Prefix)	 = strTablePrefix
		NewColumns2(22, FieldName) = "C_STROCCUPATION"
		NewColumns2(22, TableName) = "CONFIG"
		NewColumns2(22, DataType_Access) = "SMALLINT"
		NewColumns2(22, DataType_SQL6) = "SMALLINT"
		NewColumns2(22, DataType_SQL7) = "SMALLINT"
		NewColumns2(22, DataType_MySQL) = "SMALLINT"
		NewColumns2(22, ConstraintAccess)  = "NULL"
		NewColumns2(22, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1007 DEFAULT 0"
		NewColumns2(22, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1007 DEFAULT 0"
		NewColumns2(22, ConstraintMySQL)  = "DEFAULT 0 NULL"

		NewColumns2(23, Prefix)	 = strTablePrefix
		NewColumns2(23, FieldName) = "C_STRBIO"
		NewColumns2(23, TableName) = "CONFIG"
		NewColumns2(23, DataType_Access) = "SMALLINT"
		NewColumns2(23, DataType_SQL6) = "SMALLINT"
		NewColumns2(23, DataType_SQL7) = "SMALLINT"
		NewColumns2(23, DataType_MySQL) = "SMALLINT"
		NewColumns2(23, ConstraintAccess)  = "NULL"
		NewColumns2(23, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1008 DEFAULT 0"
		NewColumns2(23, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1008 DEFAULT 0"
		NewColumns2(23, ConstraintMySQL)  = "DEFAULT 0 NULL"

		NewColumns2(24, Prefix)	 = strTablePrefix
		NewColumns2(24, FieldName) = "C_STRHOBBIES"
		NewColumns2(24, TableName) = "CONFIG"
		NewColumns2(24, DataType_Access) = "SMALLINT"
		NewColumns2(24, DataType_SQL6) = "SMALLINT"
		NewColumns2(24, DataType_SQL7) = "SMALLINT"
		NewColumns2(24, DataType_MySQL) = "SMALLINT"
		NewColumns2(24, ConstraintAccess)  = "NULL"
		NewColumns2(24, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1009 DEFAULT 0"
		NewColumns2(24, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1009 DEFAULT 0"
		NewColumns2(24, ConstraintMySQL)  = "DEFAULT 0 NULL"

		NewColumns2(25, Prefix)	 = strTablePrefix
		NewColumns2(25, FieldName) = "C_STRLNEWS"
		NewColumns2(25, TableName) = "CONFIG"
		NewColumns2(25, DataType_Access) = "SMALLINT"
		NewColumns2(25, DataType_SQL6) = "SMALLINT"
		NewColumns2(25, DataType_SQL7) = "SMALLINT"
		NewColumns2(25, DataType_MySQL) = "SMALLINT"
		NewColumns2(25, ConstraintAccess)  = "NULL"
		NewColumns2(25, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1010 DEFAULT 0"
		NewColumns2(25, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1010 DEFAULT 0"
		NewColumns2(25, ConstraintMySQL)  = "DEFAULT 0 NULL"
		
		NewColumns2(26, Prefix)	 = strTablePrefix
		NewColumns2(26, FieldName) = "C_STRQUOTE"
		NewColumns2(26, TableName) = "CONFIG"
		NewColumns2(26, DataType_Access) = "SMALLINT"
		NewColumns2(26, DataType_SQL6) = "SMALLINT"
		NewColumns2(26, DataType_SQL7) = "SMALLINT"
		NewColumns2(26, DataType_MySQL) = "SMALLINT"
		NewColumns2(26, ConstraintAccess)  = "NULL"
		NewColumns2(26, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1011 DEFAULT 0"
		NewColumns2(26, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1011 DEFAULT 0"
		NewColumns2(26, ConstraintMySQL)  = "DEFAULT 0 NULL"

		NewColumns2(27, Prefix)	 = strTablePrefix
		NewColumns2(27, FieldName) = "C_STRMARSTATUS"
		NewColumns2(27, TableName) = "CONFIG"
		NewColumns2(27, DataType_Access) = "SMALLINT"
		NewColumns2(27, DataType_SQL6) = "SMALLINT"
		NewColumns2(27, DataType_SQL7) = "SMALLINT"
		NewColumns2(27, DataType_MySQL) = "SMALLINT"
		NewColumns2(27, ConstraintAccess)  = "NULL"
		NewColumns2(27, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1012 DEFAULT 0"
		NewColumns2(27, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1012 DEFAULT 0"
		NewColumns2(27, ConstraintMySQL)  = "DEFAULT 0 NULL"
		
		NewColumns2(28, Prefix)	 = strTablePrefix
		NewColumns2(28, FieldName) = "C_STRFAVLINKS"
		NewColumns2(28, TableName) = "CONFIG"
		NewColumns2(28, DataType_Access) = "SMALLINT"
		NewColumns2(28, DataType_SQL6) = "SMALLINT"
		NewColumns2(28, DataType_SQL7) = "SMALLINT"
		NewColumns2(28, DataType_MySQL) = "SMALLINT"
		NewColumns2(28, ConstraintAccess)  = "NULL"
		NewColumns2(28, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1013 DEFAULT 0"
		NewColumns2(28, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1013 DEFAULT 0"
		NewColumns2(28, ConstraintMySQL)  = "DEFAULT 0 NULL"

		NewColumns2(29, Prefix)	 = strTablePrefix
		NewColumns2(29, FieldName) = "C_STRRECENTTOPICS"
		NewColumns2(29, TableName) = "CONFIG"
		NewColumns2(29, DataType_Access) = "SMALLINT"
		NewColumns2(29, DataType_SQL6) = "SMALLINT"
		NewColumns2(29, DataType_SQL7) = "SMALLINT"
		NewColumns2(29, DataType_MySQL) = "SMALLINT"
		NewColumns2(29, ConstraintAccess)  = "NULL"
		NewColumns2(29, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1014 DEFAULT 0"
		NewColumns2(29, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1014 DEFAULT 0"
		NewColumns2(29, ConstraintMySQL)  = "DEFAULT 0 NULL"
		

		call AddColumns(NewColumns2, intCriticalErrors, intWarnings)
	
		'## Forum_SQL
		strSql = "UPDATE " & strTablePrefix & "CONFIG "
		strSql = strSql & " SET C_STRVERSION =  '" & strNewVersion & "'"
		strSql = strSql & " , C_STRFULLNAME   =  " & 0
		strSql = strSql & " , C_STRPICTURE    =  " & 0
		strSql = strSql & " , C_STRSEX        =  " & 0
		strSql = strSql & " , C_STRCITY       =  " & 0
		strSql = strSql & " , C_STRSTATE      =  " & 0
		strSql = strSql & " , C_STRAGE        =  " & 0
		strSql = strSql & " , C_STRCOUNTRY    =  " & 1
		strSql = strSql & " , C_STROCCUPATION =  " & 0
		strSql = strSql & " , C_STRHOMEPAGE   =  " & 1
		strSql = strSql & " , C_STRFAVLINKS   =  " & 1
		strSql = strSql & " , C_STRICQ        =  " & 1
		strSql = strSql & " , C_STRYAHOO      =  " & 1
		strSql = strSql & " , C_STRAIM        =  " & 1
		strSql = strSql & " , C_STRBIO        =  " & 0
		strSql = strSql & " , C_STRHOBBIES	 =  " & 0
		strSql = strSql & " , C_STRLNEWS      =  " & 0
		strSql = strSql & " , C_STRQUOTE      =  " & 0
		strSql = strSql & " , C_STRMARSTATUS  =  " & 0
		strSql = strSql & " , C_STRRECENTTOPICS  =  " & 0
		strSql = strSql & " WHERE CONFIG_ID = " & 1

		on error resume next
		my_Conn.Errors.Clear
		my_Conn.Execute (strSql)
	
		Response.Write("<table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""50%"" align=""center"">")
	
		UpdateErrorCode = UpdateErrorCheck()
	
		on error goto 0
	
		if UpdateErrorCode = 0 then
			Response.Write("<tr><td bgColor=green align=""left"" width=""30%""><font face=""宋体, Arial, Helvetica"" size=""2""><b>Upgraded: </b></font></td><td bgColor=navyblue align=""left""><font face=""宋体, Arial, Helvetica"" size=""2""> Added default values for new fields in CONFIG table</font></td></tr>")
		elseif UpdateErrorCode = 2 then
			Response.Write("<tr><td bgColor=red align=""left"" width=""30%""><font face=""宋体, Arial, Helvetica"" size=""2""><b>Critical error: </b></font></td><td bgColor=navyblue align=""left""><font face=""宋体, Arial, Helvetica"" size=""2"">Can't add default values for new fields in CONFIG table!<b></font></td></tr>")
			intCriticalErrors = intCriticalErrors + 1
		else
			Response.Write("<tr><td bgColor=red align=""left"" width=""30%""><font face=""宋体, Arial, Helvetica"" size=""2""><b>Critical error: code: </b></font></td><td bgColor=navyblue align=""left""><font face=""宋体, Arial, Helvetica"" size=""2"">" & UpdateErrorCode & " while trying to add default values to the CONFIG table</font></td></tr>")
			intCriticalErrors = intCriticalErrors + 1
		end if
		Response.Write("</table>")
		Response.Flush

	end if

'###########################################################################
'## 
'## Set up for update 3
'## 
'## 	Database updates needed
'## 
'## 	FORUM_FORUM
'## 	Need F_PASSWORD set to 255 Char's to handle NT Group Names.
'## 
'###########################################################################

	if (OldVersion <= 3) then

		Dim NewColumns3(2,11)
			
		NewColumns3(0, Prefix)	 = strTablePrefix
		NewColumns3(0, FieldName) = "C_STRAUTOLOGON"
		NewColumns3(0, TableName) = "CONFIG"
		NewColumns3(0, DataType_Access) = "SMALLINT"
		NewColumns3(0, DataType_SQL6) = "SMALLINT"
		NewColumns3(0, DataType_SQL7) = "SMALLINT"
		NewColumns3(0, DataType_MySQL) = "SMALLINT"
		NewColumns3(0, ConstraintAccess)  = "NULL"
		NewColumns3(0, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1015 DEFAULT 0"
		NewColumns3(0, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1015 DEFAULT 0"
		NewColumns3(0, ConstraintMySQL)  = "DEFAULT '0' NULL"

		NewColumns3(1, Prefix)	 = strTablePrefix
		NewColumns3(1, FieldName) = "C_STRNTGROUPS"
		NewColumns3(1, TableName) = "CONFIG"
		NewColumns3(1, DataType_Access) = "SMALLINT"
		NewColumns3(1, DataType_SQL6) = "SMALLINT"
		NewColumns3(1, DataType_SQL7) = "SMALLINT"
		NewColumns3(1, DataType_MySQL) = "SMALLINT"
		NewColumns3(1, ConstraintAccess)  = "NULL"
		NewColumns3(1, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1016 DEFAULT 0"
		NewColumns3(1, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1016 DEFAULT 0"
		NewColumns3(1, ConstraintMySQL)  = "DEFAULT '0' NULL"

		NewColumns3(2, Prefix)	 = strTablePrefix
		NewColumns3(2, FieldName) = "F_PASSWORD_NEW"
		NewColumns3(2, TableName) = "FORUM"
		NewColumns3(2, DataType_Access)  = "TEXT (255)"
		NewColumns3(2, DataType_SQL6)  = "VARCHAR (255)"
		NewColumns3(2, DataType_SQL7)  = "NVARCHAR (255)"
		NewColumns3(2, DataType_MYSQL)  = "VARCHAR (255)"
		NewColumns3(2, ConstraintAccess)  = "NULL"
		NewColumns3(2, ConstraintSQL6)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1017 DEFAULT ''"
		NewColumns3(2, ConstraintSQL7)  = "NULL CONSTRAINT " & strTablePrefix & "SnitzC1017 DEFAULT ''"
		NewColumns3(2, ConstraintMySQL)  = "DEFAULT '' NULL"

		call AddColumns(NewColumns3, intCriticalErrors, intWarnings)
		
		Dim SpecialSql3(4)
		
		SpecialSql3(Access) = "UPDATE " & strTablePrefix & "FORUM SET F_PASSWORD_NEW = F_PASSWORD"
		SpecialSql3(SQL6) = "UPDATE " & strTablePrefix & "FORUM SET F_PASSWORD_NEW = F_PASSWORD"
		SpecialSql3(SQL7) = "UPDATE " & strTablePrefix & "FORUM SET F_PASSWORD_NEW = F_PASSWORD"
		SpecialSql3(MySql) = "UPDATE " & strTablePrefix & "FORUM SET F_PASSWORD_NEW = F_PASSWORD"
		strOkMessage = "Password field conversion step 1 of 2 completed"
	
		call SpecialUpdates(SpecialSql3, strOkMessage)
		
		SpecialSql3(Access) = "ALTER TABLE " & strTablePrefix & "FORUM DROP COLUMN F_PASSWORD"
		SpecialSql3(SQL6) = "SELECT * FROM " & strTablePrefix & "CONFIG " '## dummy sql-statement SQL6.5 doesn't allow DROP !!
		SpecialSql3(SQL7) = "ALTER TABLE " & strTablePrefix & "FORUM DROP COLUMN F_PASSWORD"
		SpecialSql3(MySql) = "ALTER TABLE " & strTablePrefix & "FORUM DROP COLUMN F_PASSWORD"
		strOkMessage = "Password field conversion step 2 of 2 completed"
	
		call SpecialUpdates(SpecialSql3, strOkMessage)
		
		'## Forum_SQL
		strSql = "UPDATE " & strTablePrefix & "CONFIG "
		strSql = strSql & " SET C_STRVERSION =  '" & strNewVersion & "'"
		strSql = strSql & " , C_STRAUTOLOGON  =  " & 0
		strSql = strSql & " , C_STRNTGROUPS   =  " & 0
		strSql = strSql & " WHERE CONFIG_ID  =  " & 1

		on error resume next
		my_Conn.Errors.Clear
		my_Conn.Execute (strSql)
	
		Response.Write("<table border=0 cellspacing=0 cellpadding=5 width=""50%"" align=center>")
	
		UpdateErrorCode = UpdateErrorCheck()
	
		on error goto 0
	
		if UpdateErrorCode = 0 then
			Response.Write("<tr><td bgColor=green align=""left"" width=""30%""><font face=""宋体, Arial, Helvetica"" size=""2""><b>Upgraded: </b></font></td><td bgColor=navyblue align=""left""><font face=""宋体, Arial, Helvetica"" size=""2""> Added default values for new fields in CONFIG table</font></td></tr>")
		elseif UpdateErrorCode = 2 then
			Response.Write("<tr><td bgColor=red align=""left"" width=""30%""><font face=""宋体, Arial, Helvetica"" size=""2""><b>Critical error: </b></font></td><td bgColor=navyblue align=""left""><font face=""宋体, Arial, Helvetica"" size=""2"">Can't add default values for new fields in CONFIG table!<b></font></td></tr>")
			intCriticalErrors = intCriticalErrors + 1
		else
			Response.Write("<tr><td bgColor=red align=""left"" width=""30%""><font face=""宋体, Arial, Helvetica"" size=""2""><b>Critical error: code: </b></font></td><td bgColor=navyblue align=""left""><font face=""宋体, Arial, Helvetica"" size=""2"">" & UpdateErrorCode & " while trying to add default values to the CONFIG table</font></td></tr>")
			intCriticalErrors = intCriticalErrors + 1
		end if
		Response.Write("</table>")
		Response.Flush

	end if

'###########################################################################
'## 
'## Setup for update 4
'## 
'###########################################################################

	if (OldVersion <= 4) then
	
		Dim SpecialSql4(4)
		
		SpecialSql4(Access) = "CREATE TABLE " & strTablePrefix & "ALLOWED_MEMBERS ("
 		SpecialSql4(Access) = SpecialSql4(Access) & "MEMBER_ID INT NOT NULL, FORUM_ID INT NOT NULL, "
 		SpecialSql4(Access) = SpecialSql4(Access) & "CONSTRAINT " & strTablePrefix & "SnitzC373 PRIMARY KEY (MEMBER_ID, FORUM_ID) ) "
 	
		SpecialSql4(SQL6) = "CREATE TABLE " & strTablePrefix & "ALLOWED_MEMBERS ("
		SpecialSql4(SQL6) = SpecialSql4(SQL6) & "MEMBER_ID INT NOT NULL, FORUM_ID INT NOT NULL , "
		SpecialSql4(SQL6) = SpecialSql4(SQL6) & "CONSTRAINT " & strTablePrefix & "SnitzC373 PRIMARY KEY NONCLUSTERED (MEMBER_ID, FORUM_ID) )"
		
		SpecialSql4(SQL7) = "CREATE TABLE " & strTablePrefix & "ALLOWED_MEMBERS ("
		SpecialSql4(SQL7) = SpecialSql4(SQL7) & "MEMBER_ID INT NOT NULL, FORUM_ID INT NOT NULL , "
		SpecialSql4(SQL7) = SpecialSql4(SQL7) & "CONSTRAINT " & strTablePrefix & "SnitzC373 PRIMARY KEY NONCLUSTERED (MEMBER_ID, FORUM_ID) )"
		
		SpecialSql4(MySql) = "CREATE TABLE " & strTablePrefix & "ALLOWED_MEMBERS ("
		SpecialSql4(MySql) = SpecialSql4(MySql) & "MEMBER_ID INT (11) NOT NULL, FORUM_ID smallint (6) NOT NULL , "
		SpecialSql4(MySql) = SpecialSql4(MySql) & "PRIMARY KEY (MEMBER_ID, FORUM_ID) ) "

 		strOkMessage = "Table ALLOWED_MEMBER created "
	
		call SpecialUpdates(SpecialSql4, strOkMessage)

		Response.Flush
		
		Response.Write("<table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""50%"" align=""center"">")
		Response.Write("<tr><td bgColor=green align=""left"" width=""30%""><font face=""宋体, Arial, Helvetica"" size=""2""><b>Upgrading: </b></font></td>")
		Response.Write("<td bgColor=navyblue align=""left""><font face=""宋体, Arial, Helvetica"" size=""2"">Starting transferring Member List to ALLOWED_MEMBERS table</font></td></tr>")
		
		intTransferErrors = 0
		strSql = "SELECT FORUM_ID,F_USERLIST FROM " & strTablePrefix & "FORUM "
		
		on error resume next
		my_Conn.Errors.Clear
		
		set rsForum = my_Conn.execute(strSql)
		
		UpdateErrorCode = UpdateErrorCheck()
	
		on error goto 0
		
		if UpdateErrorCode = 0 then
			Response.Write("<tr><td bgColor=green align=""left"" width=""30%""><font face=""宋体, Arial, Helvetica"" size=""2""><b>Upgraded: </b></font></td><td bgColor=navyblue align=""left""><font face=""宋体, Arial, Helvetica"" size=""2""> Table opened</font></td></tr>")
		else
			Response.Write("<tr><td bgColor=red align=""left"" width=""30%""><font face=""宋体, Arial, Helvetica"" size=""2""><b>Critical error: </b></font></td><td bgColor=navyblue align=""left""><font face=""宋体, Arial, Helvetica"" size=""2"">Error while getting Memberlist for transfer<b></font></td></tr>")
			intCriticalErrors = intCriticalErrors + 1
			intTransferErrors = 1
		end if
		
		if intTransferErrors = 0 then
			do while not rsForum.EOF
				if Instr(rsForum("F_USERLIST"),",") > 0 then
					Users = split(rsForum("F_USERLIST"),",")
						for count = Lbound(Users) to Ubound(Users)
							strSql = "INSERT INTO " & strMemberTablePrefix & "ALLOWED_MEMBERS ("
							strSql = strSql & " MEMBER_ID, FORUM_ID) VALUES ( "& Users(count) & ", " & rsForum("FORUM_ID") & ")"	

							on error resume next
							my_Conn.Errors.Clear
							my_Conn.Execute (strSql)
						
							UpdateErrorCode = UpdateErrorCheck()
		
							on error goto 0
							if UpdateErrorCode = 0 then
								Response.Write("<tr><td bgColor=green align=""left"" width=""30%""><font face=""宋体, Arial, Helvetica"" size=""2""><b>Upgraded: </b></font></td><td bgColor=navyblue align=""left""><font face=""宋体, Arial, Helvetica"" size=""2""> Added default values for new fields in CONFIG table</font></td></tr>")
							elseif UpdateErrorCode = 2 then
								Response.Write("<tr><td bgColor=red align=""left"" width=""30%""><font face=""宋体, Arial, Helvetica"" size=""2""><b>Critical error: </b></font></td><td bgColor=navyblue align=""left""><font face=""宋体, Arial, Helvetica"" size=""2"">Error while adding record ( " & Users(count) & ", " & rsForum("FORUM_ID") & ") to ALLOWED_MEMBER table!<b></font></td></tr>")
								intCriticalErrors = intCriticalErrors + 1
								infTransferErrors = 1
							end if
						next
				end if
	 			rsForum.movenext
			loop
			
		end if
		
		on error resume next
			rsForum.close
			set rsForum = nothing
		on error goto 0
		
		
		if intTransferErrors = 0 then
			Response.Write("<tr><td bgColor=green align=""left"" width=""30%""><font face=""宋体, Arial, Helvetica"" size=""2""><b>Upgrading: </b></font></td>")
			Response.Write("<td bgColor=navyblue align=""left""><font face=""宋体, Arial, Helvetica"" size=""2"">Finished transferring Member List to ALLOWED_MEMBERS table</font></td></tr>")
		else
			Response.Write("<tr><td bgColor=red align=""left"" width=""30%""><font face=""宋体, Arial, Helvetica"" size=""2""><b>Upgrading: </b></font></td>")
			Response.Write("<td bgColor=navyblue align=""left""><font face=""宋体, Arial, Helvetica"" size=""2"">Transferring of Member List to ALLOWED_MEMBERS table was NOT succesfull !</font></td></tr>")
		end if
		
		'## Forum_SQL
		strSql = "UPDATE " & strTablePrefix & "CONFIG "
		strSql = strSql & " SET C_STRVERSION =  '" & strNewVersion & "'"
		strSql = strSql & " WHERE CONFIG_ID  =  " & 1

		on error resume next
		my_Conn.Errors.Clear
		my_Conn.Execute (strSql)
	
		Response.Write("<table border=0 cellspacing=0 cellpadding=5 width=""50%"" align=center>")
	
		UpdateErrorCode = UpdateErrorCheck()
	
		on error goto 0
	
		if UpdateErrorCode = 0 then
			Response.Write("<tr><td bgColor=green align=""left"" width=""30%""><font face=""宋体, Arial, Helvetica"" size=""2""><b>Upgraded: </b></font></td><td bgColor=navyblue align=""left""><font face=""宋体, Arial, Helvetica"" size=""2""> Added default values for new fields in CONFIG table</font></td></tr>")
		elseif UpdateErrorCode = 2 then
			Response.Write("<tr><td bgColor=red align=""left"" width=""30%""><font face=""宋体, Arial, Helvetica"" size=""2""><b>Critical error: </b></font></td><td bgColor=navyblue align=""left""><font face=""宋体, Arial, Helvetica"" size=""2"">Can't add default values for new fields in CONFIG table!<b></font></td></tr>")
			intCriticalErrors = intCriticalErrors + 1
		else
			Response.Write("<tr><td bgColor=red align=""left"" width=""30%""><font face=""宋体, Arial, Helvetica"" size=""2""><b>Critical error: code: </b></font></td><td bgColor=navyblue align=""left""><font face=""宋体, Arial, Helvetica"" size=""2"">" & UpdateErrorCode & " while trying to add default values to the CONFIG table</font></td></tr>")
			intCriticalErrors = intCriticalErrors + 1
		end if
		Response.Write("</table>")
		Response.Flush
		
		
		
	end if

	if OldVersion = 5 then
	
	'## just update the version info...
	
		strSql = "UPDATE " & strTablePrefix & "CONFIG "
		strSql = strSql & " SET C_STRVERSION =  '" & strNewVersion & "'"
		strSql = strSql & " WHERE CONFIG_ID = " & 1

		on error resume next
		my_Conn.Errors.Clear
		my_Conn.Execute (strSql)
	
		Response.Write("<table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""50%"" align=""center"">")
	
		UpdateErrorCode = UpdateErrorCheck()
	
		on error goto 0
	
		if UpdateErrorCode = 0 then
			Response.Write("<tr><td bgColor=green align=""left"" width=""30%""><font face=""宋体, Arial, Helvetica"" size=""2""><b>Upgraded: </b></font></td><td bgColor=navyblue align=""left""><font face=""宋体, Arial, Helvetica"" size=""2""> Added default values for new fields in CONFIG table</font></td></tr>")
		elseif UpdateErrorCode = 2 then
			Response.Write("<tr><td bgColor=red align=""left"" width=""30%""><font face=""宋体, Arial, Helvetica"" size=""2""><b>Critical error: </b></font></td><td bgColor=navyblue align=""left""><font face=""宋体, Arial, Helvetica"" size=""2"">Can't add default values for new fields in CONFIG table!<b></font></td></tr>")
			intCriticalErrors = intCriticalErrors + 1
		else
			Response.Write("<tr><td bgColor=red align=""left"" width=""30%""><font face=""宋体, Arial, Helvetica"" size=""2""><b>Critical error: code: </b></font></td><td bgColor=navyblue align=""left""><font face=""宋体, Arial, Helvetica"" size=""2"">" & UpdateErrorCode & " while trying to add default values to the CONFIG table</font></td></tr>")
			intCriticalErrors = intCriticalErrors + 1
		end if
		Response.Write("</table>")
		Response.Flush
	
	end if

my_Conn.Close
set my_Conn = nothing

if intCriticalErrors = 0 then
%>
	<p><font face="宋体, Arial, Helvetica" size="4">The Upgrade has been completed without errors !</font></p>
<%
else
%>
	<p><font face="宋体, Arial, Helvetica" size="4"><b>The Upgrade has NOT been completed without errors !</b></font></p>
	<p><font face="宋体, Arial, Helvetica" size="2">There were <%=intCriticalErrors %>  Critical Errors...</font></p>
<%
	if intWarnings > 0 then
%>
	<p><font face="宋体, Arial, Helvetica" size="2">There were <%=intWarnings %>  noncritical errors...</font></p>
<%
	end if
end if
if intCriticalErrors > 0 then
%>
	<p><font face="宋体, Arial, Helvetica" size="2"><a href="setup.asp?RC=3">Click here to retry....</a></font></p>
<%
end if
%>
	<p><font face="宋体, Arial, Helvetica" size="2"><a href="default.asp" target="_top">Click here to return to the forum....</a></font></p>
	
</center></div>
<%
		Application(strCookieURL & "ConfigLoaded")= ""

	else
		Response.Redirect "setup_logon.asp"
	end if

elseif ResponseCode = 5 then '## install a new database

	if strDBType = "access" then '## install is not available for Access !
		Response.Redirect "setup.asp"
	else

%>	
	<div align=center><center>
	<p><font face="宋体, Arial, Helvetica" size="4">Installation of forum-tables in the database.</font></p>
	</center></div>
<form action="setup.asp?RC=6" method="post" id=Form1 name=Form1>
<table border="0" cellspacing="0" cellpadding="5" width="50%" height="50%" align="center">
	<tr>
		<td bgColor=navyblue align="left"><font face="宋体, Arial, Helvetica" size="2">
		<p>
		<% if strDBType = "sqlserver" then %>
		  <p>Select the SQL-server version you are using:</p>
          <p><input type="radio" class="radio" name="SQL_Server" value="SQL6" >SQL-Server 6.5<br>
          <input type="radio" class="radio" checked name="SQL_Server" value="SQL7">SQL-Server 7&nbsp;&nbsp;&nbsp;</p>
        <% end if %>
          <p>To install the tables in the database you need to create the empty database on the server first.<br>
          Then you have to provide a username and password of a user that has table creation/modification rights at the database you use.<br>
          This might not be the same user as you use in your connectionstring !<br>
          <br>
          Name:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input type="text" name="DBUserName" size="20"><br>
          Password: <input type="password" name="DBPassword" size="20"></p>
   
          </font>
        </td>
		</tr>	
		<tr>
			  <td align="center"><input type="submit" value="Continue" id="Submit1" name="Submit1"></td>
		</tr>
		</form>
</table>	
<%
	end if

elseif ResponseCode = 6 then '## start installing the tables in the database

	if strDBType = "access" Then '## install is not available for access !
		Response.Redirect "setup.asp"
	else

%>	
	<div align=center><center>
	<p><font face="宋体, Arial, Helvetica" size="4">Please Wait until the installation has been completed !</font></p>
<%
	on error resume next
	
	strInstallString = strConnString
	
	if strDBType <> "access" then
		strInstallString = CreateConnectionString(strConnString, Request.Form("DBUserName"), Request.Form("DBPassword"))
	end if
	
	set my_Conn = Server.CreateObject("ADODB.Connection")
	my_Conn.Open strInstallString
	
	for counter = 0 to my_conn.Errors.Count -1
		ConnErrorNumber = my_conn.Errors(counter).Number
		if ConnErrorNumber <> 0 then 
			my_conn.Errors.Clear 
			Response.Redirect "setup.asp?RC=1&EC=" & ConnErrorNumber & "&RET=" & Server.URLEncode("setup.asp?RC=5")
		end if
	next
	
	on error goto 0
	
	intCriticalErrors = 0
	
%>
<table border="0" cellspacing="0" cellpadding="5" width="50%" height="50%" align="center">
	<tr>
	<%
strSQL_Server = Request.Form("Sql_Server")

if strDBType = "mysql" then %>
<!--#INCLUDE FILE="inc_create_forum_mysql.asp" -->
<%
elseif strDBType = "sqlserver" then
	if strSQL_Server = "SQL6" then %>
<!--#INCLUDE FILE="inc_create_forum_sql6.asp" -->
<%  else
%>
<!--#INCLUDE FILE="inc_create_forum_sql7.asp" -->
<%  end if
end if %>
	</tr>
</table>
<%

	my_Conn.Close
	set my_Conn = nothing
	

	if intCriticalErrors = 0 then
%>
	<p><font face="宋体, Arial, Helvetica" size="4">The Installation has been completed !</font></p>
<%
	else
%>
	<p><font face="宋体, Arial, Helvetica" size="4"><b>The Installation has NOT been completed !</b></font></p>
	<p><font face="宋体, Arial, Helvetica" size="2">There were <%=intCriticalErrors %>  Critical Errors...</font></p>
<%
	end if
	if intCriticalErrors > 0 then
%>
	<p><font face="宋体, Arial, Helvetica" size="2"><a href="setup.asp?RC=5">Click here to retry....</a></font></p>
<%
	end if
%>
	<p><font face="宋体, Arial, Helvetica" size="2"><a href="setup.asp" target="_top">Click here to check the Database....</a></font></p>
	
</center></div>

<%
	end if
else
%>
<html>

<head>
<title>Undefined error !!</title>
<meta name="copyright" content="Snitz Forums 2000 Version 3.1 SR4 汉化修改：资源搜罗站">
<Style><!--
a:link    {color:<% =strLinkColor %>;text-decoration:<% =strLinkTextDecoration %>}
a:visited {color:<% =strVisitedLinkColor %>;text-decoration:<% =strVisitedTextDecoration %>}
a:hover   {color:<% =strHoverFontColor %>;text-decoration:<% =strHoverTextDecoration %>}
--></style>
</head>

<body bgColor="white" text="midnightblue" link="darkblue" aLink=red vLink="red" onLoad="window.focus()">


<div align=center><center>
	<font face="宋体, Arial, Helvetica" size="4">There has been an error !!</font>
</center></div>

<table border="0" cellspacing="0" cellpadding="5" width="100%" align="left">
	<tr>
		<td bgColor=navyblue align="center">
		<font face="宋体, Arial, Helvetica" size="2">
<%		Response.Write(Err.number & ", " & Err.description) %>
		</font></td>
	</tr>
	<tr>
		<td align="center">
		<font face="宋体, Arial, Helvetica" size="2">
		<a href="default.asp" target="_top">Click here to retry.</a>
		</font></td>
	</tr>
</table>	
<%
End if
%>

</body>
</html>

<%
sub CheckSqlError()

dim ChkConnErrorNumber

	for counter = 0 to my_Conn.Errors.Count -1
		ChkConnErrorNumber = my_Conn.Errors(counter).Number
		if  ChkConnErrorNumber <> 0 or Err.number > 0 then  
		
			my_Conn.Errors.Clear 
			Err.Clear
		
			strSql = "SELECT " & strTablePrefix & "CONFIG.C_STRVERSION, "
			strSql = strSql & strTablePrefix & "CONFIG.C_STRSENDER "
			strSql = strSql & " FROM " & strTablePrefix & "CONFIG "

			set rsInfo = my_Conn.Execute (StrSql)
			strVersion = rsInfo("C_STRVERSION")
			strSender = rsInfo("C_STRSENDER")
		
			rsInfo.Close
			set rsInfo = nothing
			my_Conn.Close
			set my_Conn = nothing
		
			Response.Redirect "setup.asp?RC=2&MAIL=" & Server.UrlEncode(strSender) & "&VER=" & Server.URLEncode(strVersion) & "&EC=" & ChkConnErrorNumber
		end if
	next
end sub

function UpdateErrorCheck()

dim intErrorNumber
dim counter

	intErrorNumber = 0
	for counter = 0 to my_Conn.Errors.Count -1
		intErrorNumber = my_Conn.Errors(counter).Number
		if intErrorNumber <> 0 then  
			select case intErrorNumber
				case -2147217900
					UpdateErrorCheck = 1
					counter = my_Conn.Errors.Count -1
				case -2147467259
					UpdateErrorCheck = 2
					counter = my_Conn.Errors.Count -1	
				case else
					UpdateErrorCheck = intErrorNumber
			end select
		end if
	next
end function

Sub AddColumns(Columns, intCriticalErrors, intWarnings)

Dim colCounter

	Response.Write("<table border=""0"" cellspacing=""1"" cellpadding=""5"" width=""50%"" height=""50%"" align=""center"">")
	For colCounter = 0 to Ubound(Columns, 1) 
		on error resume next
		my_Conn.Errors.Clear
		
		strUpdateSql = "ALTER TABLE " & Columns(colCounter, Prefix) & Columns(colCounter, TableName) & "  "
		strUpdateSql = strUpdateSql & " ADD " & Columns(colCounter, FieldName) & " " 
		if strDBType = "access" then
			strUpdateSql = strUpdateSql & " " & Columns(colCounter, DataType_Access) & " " & Columns(colCounter, ConstraintAccess) & " ; "
		elseif strDBType = "sqlserver" then 
			if strSQL_Server = "SQL7" then
				strUpdateSql = strUpdateSql & " " & Columns(colCounter, DataType_SQL7) & " " & Columns(colCounter, ConstraintSQL7) & " ; "
			else
				strUpdateSql = strUpdateSql & " " & Columns(colCounter, DataType_SQL6) & " " & Columns(colCounter, ConstraintSQL6) & " ; "
			end if
		elseif strDBType = "mysql" then
			strUpdateSql = strUpdateSql & "  " & Columns(colCounter, DataType_MySql) & " " & Columns(colCounter, ConstraintMySql) & " "
		end if
								
		my_Conn.Execute strUpdateSql
	
		UpdateErrorCode = UpdateErrorCheck()
		
		on error goto 0
		
		if UpdateErrorCode = 0 then
			Response.Write("<tr><td bgColor=green align=""left"" width=""30%""><font face=""宋体, Arial, Helvetica"" size=""2""><b>Upgraded:</b></font></td><td bgColor=navyblue align=""left""><font face=""宋体, Arial, Helvetica"" size=""2"">" & Columns(colCounter, FieldName) & " has been added to the " & Columns(colCounter, TableName) & " table</font></td></tr>")
		elseif UpdateErrorCode = 1 then
			Response.Write("<tr><td bgColor=orange align=""left"" width=""30%""><font face=""宋体, Arial, Helvetica"" size=""2""><b>Noncritical error: </b></font></td><td bgColor=navyblue align=""left""><font face=""宋体, Arial, Helvetica"" size=""2"">" & Columns(colCounter, Fieldname) & " allready existed in the " & Columns(colCounter, TableName) & " table</font></td></tr>")
			intWarnings = intWarnings + 1
		elseif UpdateErrorCode = 2 then
			Response.Write("<tr><td bgColor=red align=""left"" width=""30%""><font face=""宋体, Arial, Helvetica"" size=""2""><b>Critical error: </b></font></td><td bgColor=navyblue align=""left""><font face=""宋体, Arial, Helvetica"" size=""2""> No write access to the table " & Columns(colCounter, TableName) & " <br>" & Columns(colCounter, Fieldname) & " not added to database!</font></td></tr>")
			intCriticalErrors = intCriticalErrors + 1
		else
			Response.Write("<tr><td bgColor=red align=""left"" width=""30%""><font face=""宋体, Arial, Helvetica"" size=""2""><b>Critical error: code: </b></font></td><td bgColor=navyblue align=""left""><font face=""宋体, Arial, Helvetica"" size=""2"">" & UpdateErrorCode & " in statement [" & strUpdateSql & "] while trying to add " & Columns(colCounter, Fieldname) & " to the " & Columns(colCounter, TableName) & " table</font></td></tr>")
			intCriticalErrors = intCriticalErrors + 1
		end if
		Response.Flush
	Next
	Response.Write("</table>")
end sub

sub SpecialUpdates(strSql, strOkMessage)

dim strUpdateSql
dim SpecialErrors

	on error resume next
	my_Conn.Errors.Clear
		
	if strDBType = "access" then
		strUpdateSql = strSql(Access) 
	elseif strDBType = "sqlserver" then 
		if strSQL_Server = "SQL7" then
			strUpdateSql = strSql(SQL7)
		else
			strUpdateSql = strSql(SQL6)
		end if
	elseif strDBType = "mysql" then
		strUpdateSql = strSql(MySql)
	end if
								
	my_Conn.Execute strUpdateSql
	
	SpecialErrors = 0
	Response.Write("<table border=""0"" cellspacing=""1"" cellpadding=""5"" width=""50%"" align=""center"">")	
	for counter = 0 to my_conn.Errors.Count -1
		ConnErrorNumber = my_conn.Errors(counter).Number
		ConnErrorDescription = my_conn.Errors(counter).Description
		
		if ConnErrorNumber <> 0 then 
		
			if ConnErrorNumber = -2147217900 then
				Response.Write("<tr><td bgColor=orange align=""left"" width=""30%""><font face=""宋体, Arial, Helvetica"" size=""2""><b>Error: " & ConnErrorNumber & "</b></font></td>")
				Response.Write("<td bgColor=navyblue align=""left""><font face=""宋体, Arial, Helvetica"" size=""2"">" & ConnErrorDescription & "</font></td></tr>")
				Response.Write("<tr><td bgColor=orange align=""left"" width=""30%""><font face=""宋体, Arial, Helvetica"" size=""2""><b>strUpdateSql: </b></font></td>")
				Response.Write("<td bgColor=navyblue align=""left""><font face=""宋体, Arial, Helvetica"" size=""2"">" & strUpdateSql & "</font></td></tr>")
				intWarnings = intWarnings + 1
				SpecialErrors = 1
			else
				Response.Write("<tr><td bgColor=red align=""left"" width=""30%""><font face=""宋体, Arial, Helvetica"" size=""2""><b>Error: " & ConnErrorNumber & "</b></font></td>")
				Response.Write("<td bgColor=navyblue align=""left""><font face=""宋体, Arial, Helvetica"" size=""2"">" & ConnErrorDescription & "</font></td></tr>")
				Response.Write("<tr><td bgColor=red align=""left"" width=""30%""><font face=""宋体, Arial, Helvetica"" size=""2""><b>strUpdateSql: </b></font></td>")
				Response.Write("<td bgColor=navyblue align=""left""><font face=""宋体, Arial, Helvetica"" size=""2"">" & strUpdateSql & "</font></td></tr>")
				intCriticalErrors = intCriticalErrors + 1
				SpecialErrors = 1
			end if
		end if
	next
	if SpecialErrors = 0 then
		Response.Write("<tr><td bgColor=green align=""left"" width=""30%""><font face=""宋体, Arial, Helvetica"" size=""2""><b>Upgraded: </b></font></td>")
		Response.Write("<td bgColor=navyblue align=""left""><font face=""宋体, Arial, Helvetica"" size=""2"">" & strOkMessage & "</font></td></tr>")
	end if
	
	Response.Write("</table>")
	Response.Flush
	
	my_conn.Errors.Clear 
	on error goto 0

end sub


Function CreateConnectionString(strConn, UserName, Password)
'## strConnString = "driver={SQL Server};server=YYYY;uid=XX;pwd=XXXX;database=ZZZZZZ"
'## strConnString = "driver={SQL Server};server=YYYY;USER ID=XX;PASSWORD=XXXX;database=ZZZZZZ"

Dim TempConnString
Dim uidTagStart, uidTagEnd
Dim pwdTagStart, pwdTagEnd
Dim uidTagStartPos, uidTagEndPos
Dim pwdTagStartPos, pwdTagEndPos

uidTagStart(0)		= "UID="
uidTagStart(1)		= "USER ID="
uidTagEnd			= ";" 
pwdTagStart(0)		= "PWD="
pwdTagStart(1)		= "PASSWORD="
pwdTagEnd			= ";"

TempConnString = strConn

for Counter = 0 to Ubound(uidTagStart)

	uidTagStartPos = InStr(1, UCase(TempConnString), UCase(uidTagStart(Counter)), 1)
	uidTagEndPos = InStr(uidTagStartPos, UCase(TempConnString), UCase(uidTagEnd(Counter)), 1)

	if (uidTagStartpos > 0) and (uidTagEndPos > 0) then
		TempConnString = Left(TempConnString, (uidTagStartPos + len(uidTagStart(Counter))-1)) & UserName & Right(TempConnString, (len(TempConnString) - uidTagEndPos) + 1)
	end if

	pwdTagStartPos = InStr(1, TempConnString, pwdTagStart(Counter), 1)
	pwdTagEndPos = InStr(pwdTagStartPos, TempConnString, pwdTagEnd(Counter), 1)

	if (pwdTagStartpos > 0) and (pwdTagEndPos > 0) then
		TempConnString = Left(TempConnString, (pwdTagStartPos + len(pwdTagStart(Counter))-1)) & Password & Right(TempConnString, (len(TempConnString) - pwdTagEndPos) + 1)
	end if	
	
next
	CreateConnectionString = TempConnString
end function
%>
	
