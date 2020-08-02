<!--#INCLUDE FILE="config.asp" -->
<!--#INCLUDE FILE="inc_functions.asp" -->
<!--#INCLUDE FILE="inc_top_short.asp" -->
<meta http-equiv="Refresh" content="30; URL=pop_online.asp">
<table width=100% border="0" height=100% cellspacing="0" cellpadding="0" align="center">
<tr bgcolor=#6666ff>
<td bgcolor=#6666ff>
<table width="100%" border="0" cellspacing="0" cellpadding="0" valign="top" align="center">
				<tr width=100%>
					<td width=20><img src="<%=strImageURL %>nav_top_left.gif" height=20 width=20 border=0></td>
					<td width=20><img src="<%=strImageURL %>nav_top.gif" height=20 width="20" border=0></td>
					<td width=102><img src="<%=strImageURL %>vbwhos.gif" width=102 height=20 border=0></td>
					<td width=100%><img src="<%=strImageURL %>nav_top.gif" height=20 width="100%" border=0></td>
					<td width=20><img src="<%=strImageURL %>nav_top_right.gif" height=20 width=20 border=0></td>
				</tr>
</table>
</td>
</tr>
<tr bgcolor=#6666ff>
<td bgcolor=#6666ff>
<table width="100%" border="0" cellspacing="0" cellpadding="0" valign="top" align="center">
<tr align=left>
<td>
<br>
<%
	
Dim objConn, strConn, CheckInTime, SQL, User, Timedout, Date
Dim objRS, strSQL, houron, minon, Datec
' WHOS ONLINE SCRIPT
Dim strOnlinePathInfo, strOnlineQueryString, strOnlineLocation
Dim strOnlineUser, strOnlineDate, strOnlineCheckInTime, strOnlineTimedOut
Dim strOnlineUsersCount, strOnlineGuestsCount, strOnlineMembersCount
Dim strOnlineGuestUserIP

' ******************************************************
' ADD HERE WHAT YOU WANT THE PREFIX OF YOUR COOKIE TO BE
' it will either be 'strCookieURL' or 'strUniqueID'
strTempCookieType = strUniqueID
' ******************************************************




Function OnlineSQLencode(byVal strPass)
 If not isNull(strPass) and strPass <> "" Then
 	strPass = Replace(strPass, "%", "'%'")
 	strPass = Replace(strPass, "'", "''")
 	strPass = Replace(strPass, "|", "'|'")
 	OnlineSQLencode = strPass
 End If
End Function

Function OnlineSQLdecode(byVal strPass)
 If not isNull(strPass) and strPass <> "" Then
 	strPass = Replace(strPass, "'%'", "%")
 	strPass = Replace(strPass, "''", "'")
 	strPass = Replace(strPass, "'|'", "|")
 	OnlineSQLdecode = strPass
 End If
End Function


' LETS GET WHAT PAGE THEY ARE ON
strOnlinePathInfo = Request.ServerVariables("Path_Info")
strOnlineQueryString = Request.QueryString

' TRY AND FIND OUT WHAT PAGE THEY ARE ON
If lcase(Right(strOnlinePathInfo, 9)) = "forum.asp" Then
	strOnlineLocation = "<a href=""forum.asp?" & strOnlineQueryString & """>" & Request.QueryString("Forum_Title") & "</a>"
ElseIf lcase(Right(strOnlinePathInfo, 11)) = "default.asp" Then
	strOnlineLocation = "<a href=""default.asp"">论坛首页</a>"
ElseIf lcase(Right(strOnlinePathInfo, 9)) = "topic.asp" Then
	strOnlineLocation = "查看主题' <a href=""link.asp?TOPIC_ID=" & Request.QueryString("TOPIC_ID") & """>" & Request.QueryString("Topic_Title") & "</a> '"
ElseIf lcase(Right(strOnlinePathInfo, 8)) = "post.asp" Then
	If Request.QueryString("method") = "Reply" Then
		strOnlineLocation = "回复主题' <a href=""topic.asp?" & strOnlineQueryString & """>" & ChkString(Request.QueryString("Topic_Title"),"title")  & "</a> '"
	ElseIf Request.QueryString("method") = "Topic" Then
		strOnlineLocation = "发表新主题 ' <a href=""forum.asp?" & strOnlineQueryString & """>" & Request.QueryString("Forum_Title") & "</a> '"
	Else
		strOnlineLocation = "不知道在做什么"
	End If
ElseIf lcase(Right(strOnlinePathInfo, 10)) = "active.asp" Then
	strOnlineLocation = "<a href=""active.asp"">浏览最新文章</a>"
ElseIf lcase(Right(strOnlinePathInfo, 11)) = "members.asp" Then
	strOnlineLocation = "<a href=""members.asp"">浏览论坛会员列表</a>"
ElseIf lcase(Right(strOnlinePathInfo, 10)) = "search.asp" Then
	strOnlineLocation = "<a href=""search.asp"">好像准备搜寻什么东西似的</a>"
ElseIf lcase(Right(strOnlinePathInfo, 7)) = "faq.asp" Then
	strOnlineLocation = "<a href=""faq.asp"">正在查看帮助说明</a>"
ElseIf lcase(Right(strOnlinePathInfo, 15)) = "pop_profile.asp" Then
	If Request.QueryString("mode") = "Display" Then
		strOnlineLocation = "<a href=""pop_profile.asp?" & strOnlineQueryString & """>正在个人资料栏目</a> '"
	Else
		strOnlineLocation = "个人资料"
	End If
ElseIf lcase(Right(strOnlinePathInfo, 11)) = "pm_view.asp" Then
	strOnlineLocation = "<a href=""pm_view.asp"">正在查看悄悄话收件箱</a>"
ElseIf lcase(Right(strOnlinePathInfo, 14)) = "pm_options.asp" Then
	strOnlineLocation = "<a href=""pm_view.asp"">查看悄悄话参数设置</a>"
ElseIf lcase(Right(strOnlinePathInfo, 15)) = "privatesend.asp" Then
	strOnlineLocation = "<a href=""privatesend.asp"">正在发送悄悄话</a>"
ElseIf lcase(Right(strOnlinePathInfo, 16)) = "active_users.asp" Then
	strOnlineLocation = "<a href=""active_users.asp"">正在查看在线人员列表</a>"
Else
	strOnlineLocation = "不知道在哪里"
End If

' FIND OUT IF THEY ARE A GUEST, OR A USER
if Request.Cookies(strTempCookieType & "User")("Name") = "" then
	strOnlineUser = "Guest"
else
	strOnlineUser = Request.Cookies(strTempCookieType & "User")("Name")
end if

strOnlineUserIP = remoteIP()

' LETS ENCODE THIS INFO
strOnlineUser = OnlineSQLencode(strOnlineUser)
strOnlineLocation = OnlineSQLencode(strOnlineLocation)

' SET WHEN TO TIMEOUT THE USER
' DO THIS IN SECONDS
strOnlineDate = DateToStr(Date)
strOnlineCheckInTime = DateToStr(Now())

strOnlineTimedOut = strOnlineCheckInTime - 660      'time out the user after 11 minutes ( 660 seconds )

strSql = "SELECT " & strTablePrefix & "ONLINE.UserID, " & strTablePrefix & "ONLINE.UserIP, " & strTablePrefix & "ONLINE.LastChecked"
strSql = strSql & " FROM " & strTablePrefix & "ONLINE "
strSql = strSql & " WHERE " & strTablePrefix & "ONLINE.UserIP='" & strOnlineUserIP & "' AND " & strTablePrefix & "ONLINE.UserID='" & strOnlineUser & "'"
set rsWho =  my_Conn.Execute (strSql)

if rsWho.eof or rsWho.bof then
	' THEY ARE A NEW USER SO INSERT THERE USERNAME
	on error resume next
	Set objRS2 = Server.CreateObject("ADODB.Recordset")
	strSQL =  "INSERT INTO " & strTablePrefix & "ONLINE (UserID,UserIP,DateCreated,CheckedIn,LastChecked,M_BROWSE) VALUES ('"
	strSql = strSQL & strOnlineUser & "','" & strOnlineUserIP & "','" & strOnlineDate & "','" & strOnlineCheckInTime & "','" & strOnlineCheckInTime & "','" & strOnlineLocation & "')"
	my_Conn.Execute (strSql)
	if err.number <> 0 then response.write err.number & "|" & err.description
else
	' THEY ARE A ACTIVE USER
	strSql = "SELECT " & strTablePrefix & "ONLINE.UserID, " & strTablePrefix & "ONLINE.UserIP, " & strTablePrefix & "ONLINE.LastChecked"
	strSql = strSql & " FROM " & strTablePrefix & "ONLINE "
	strSql = strSql & " WHERE " & strTablePrefix & "ONLINE.UserID='" & strOnlineUser & "' AND " & strTablePrefix & "ONLINE.UserIP = '" & strOnlineUserIP & "'"
	set rsLastChecked =  my_Conn.Execute (strSql)

	' LETS UPDATE THE TABLE SO IT SHOWS THERE LAST ACTIVE VISIT
	strSql = "UPDATE " & strTablePrefix & "ONLINE SET M_BROWSE='" & strOnlineLocation & "' , LastChecked='" & strOnlineCheckInTime & "' WHERE UserID='" & strOnlineUser & "' AND " & strTablePrefix & "ONLINE.UserIP='" & strOnlineUserIP & "'"
	my_Conn.Execute (strSql)
end if

' LETS DELETE ALL INACTIVE USERS
SQL = "DELETE FROM " & strTablePrefix & "ONLINE WHERE LastChecked < '" & strOnlineTimedOut & "'"
my_Conn.Execute SQL

set rsOnline = Server.CreateObject("ADODB.Recordset")

if strDBType = "access" then
	strSqL = "SELECT count(UserID) AS [onlinecount] "
else
	strSqL = "SELECT count(UserID) AS onlinecount  "
end if

strSql = strSql & "FROM " & strTablePrefix & "ONLINE "
Set rsOnline = my_Conn.Execute(strSql)
onlinecount = rsOnline("onlinecount")
strOnlineUsersCount = rsOnline("onlinecount")

' Get Guest count for display on Default.asp
set rsGuests = Server.CreateObject("ADODB.Recordset")
if strDBType = "access" then
	strSqL = "SELECT count(UserID) AS [Guests] "
else
	strSqL = "SELECT count(UserID) Guests  "
end if
strSql = strSql & "FROM " & strTablePrefix & "ONLINE "
strSql = strSql & " WHERE Right(UserID, 5) = 'Guest' "

Set rsGuests = my_Conn.Execute(strSql)
Guests = rsGuests("Guests")
strOnlineGuestsCount = rsGuests("Guests")


' Get Member count for display on Default.asp
set rsGuests = Server.CreateObject("ADODB.Recordset")
if strDBType = "access" then
	strSqL = "SELECT count(UserID) AS [Members] "
else
	strSqL = "SELECT count(UserID) Members  "
end if
strSql = strSql & "FROM " & strTablePrefix & "ONLINE "
strSql = strSql & " WHERE Right(UserID, 5) <> 'Guest' "

Set rsMembers = my_Conn.Execute(strSql)
Members = rsMembers("Members")
strOnlineMembersCount = rsMembers("Members")

' END WHOS ONLINE SCRIPT

	Set objRS = Server.CreateObject("ADODB.Recordset")
		
	strSql ="SELECT " & strTablePrefix & "ONLINE.UserID, " & strTablePrefix & "ONLINE.UserIP, " & strTablePrefix & "ONLINE.M_BROWSE, " & strTablePrefix & "ONLINE.DateCreated, " & strTablePrefix & "ONLINE.LastChecked, " & strTablePrefix & "ONLINE.CheckedIn "
	strSql = strSql & " FROM " & strTablePrefix & "ONLINE "
	strSql = strSql & " ORDER BY " & strTablePrefix & "ONLINE.DateCreated, " & strTablePrefix & "ONLINE.CheckedIn DESC"

	objRS.Open strSQL, my_Conn, 3
	num = 0
	If objRS.EOF or objRS.BOF Then
		response.write "<center>&nbsp;&nbsp;<font face=Arial color=yellow size=1><b>没有会员在线.</b></font></center>"
	Else
	  	While Not objRS.EOF
			Datec = objRS("DateCreated")
			if Right(objRS("UserID"), 5) = "Guest" then 
	    		num = num + 1
	    		response.write "<font face=""Arial,helvetica"" size=1 color=white><strong><font color=yellow>&nbsp;游客 #" & num & "</font></strong>"
	    
	    	else 
				response.write "<font face=""Arial,helvetica"" size=1 color=white><strong><font color=yellow>&nbsp;" & objRS("UserID") & "</font></strong>"
	   		end if
	   		response.write "&nbsp;enters @ <font color=white>&nbsp;" & ChkDate(ObjRS("CheckedIn")) & chkTime(ObjRS("CheckedIn"))  & "</font></font><br>"
			ObjRS.MoveNext
		Wend
  	End If
%>
</td>
<tr>
</table>
</td>
</tr>
<tr height=100% bgcolor=#6666ff>
<td bgcolor=#6666ff>&nbsp;
</td>
</tr>
<tr bgcolor=#6666ff>
<td bgcolor=#6666ff>
	<table bgcolor=white width="100%" border="0" cellspacing="0" cellpadding="0" align="center" valign="bottom">
		<tr>
			<td><img src="<%=strImageURL %>nav_bot_left.gif" height=20 width=20>
			</td>
			<td width=100%><img src="<%=strImageURL %>nav_bot.gif" height=20 width=100%>
			</td>
			<td><a href="javascript://" onclick="window.close(this)"><img src="<%=strImageURL %>vbclose.gif" border=0" height=20 width=78></a>
			</td>
			<td width=100%><img src="<%=strImageURL %>nav_bot.gif" height=20 width=100%>
			</td>
			<td><img src="<%=strImageURL %>nav_bot_right.gif" height=20 width=20>
			</td>
		</tr>
	</table>
</td>
</tr>
</table>
<!--#INCLUDE FILE="inc_footer_short.asp" -->

