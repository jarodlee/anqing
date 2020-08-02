<%
dim mLev, strLoginStatus,whereSql

function remoteIP()	
	
	remoteIP = Request.ServerVariables("HTTP_X_FORWARDED_FOR")	
	if remoteIP = "" then		
		remoteIP = Request.ServerVariables("REMOTE_ADDR")	
	end if
end function


set my_Conn = Server.CreateObject("ADODB.Connection")
my_Conn.Open strConnString


if (strAuthType = "nt") then
	call NTauthenticate()
	if (ChkAccountReg() = "1") then
		call NTUser()
	end if
end if

	strDBNTUserName = Request.Cookies(strUniqueID & "User")("Name")
	strDBNTFUserName = Request.Form("Name")
	if strAuthType = "nt" then
		strDBNTUserName = Session(strCookieURL & "userID")
		strDBNTFUserName = Session(strCookieURL & "userID")
	end if

	
select case Request.Form("Method_Type")
	case "login"
		select case ChkUser2(strDBNTFUserName, Request.Form("Password"))
			case 1, 2, 3, 4
				Call DoCookies(Request.Form("SavePassword"))
				strLoginStatus = 1
			case else
				strLoginStatus = 0
			end select
	case "logout"
		Call ClearCookies()
				'delete user if exists
				SQL = "DELETE FROM " & strTablePrefix & "ONLINE WHERE " & strTablePrefix & "ONLINE.UserID = '" & strDBNTUserName & "'"
				my_conn.Execute SQL
				SQL = "DELETE FROM " & strTablePrefix & "ONLINE WHERE " & strTablePrefix & "ONLINE.UserID = '" & remoteIP() & "Guest" & "'"
				my_Conn.Execute SQL
		
end select

mLev = cint(ChkUser2(strDBNTUserName, Request.Cookies(strUniqueID & "User")("Pword")))
if request.serverVariables("PATH_INFO") <> (strCookieURL & "admin_login.asp") then
if (Application("down") AND mLev <> 4) then
response.redirect "mods/down.asp"
end if
end if


strYOffset=180
strStaticYOffset=40
strMenuWidth=140

Response.Write	"<html>" & vbcrlf & vbcrlf & _
				"<head>" & vbcrlf & _
				"<title>" & strForumTitle & "</title>" & vbcrlf & _
				"<meta name=""copyright"" content=""Snitz Forums 2000 Version 3.1 SR4 资源搜罗站汉化修改(http://www.99ss.net)"">" & vbcrlf


				%>
<script language="JavaScript">
<!-- hide from JavaScript-challenged browsers
function openWindow(url) {
  popupWin = window.open(url,'new_page','width=400,height=400')
}
function openWindow2(url) {
  popupWin = window.open(url,'new_page','width=400,height=450')
}
function openWindow3(url) {
  popupWin = window.open(url,'new_page','width=400,height=450,scrollbars=yes')
}
function openWindow4(url) {
  popupWin = window.open(url,'new_page','width=400,height=525')
}
function openWindow5(url) {
popupWin = window.open(url,'new_page','width=400,height=500,scrollbars=1')
}
function openpollWindow(url) {
  popupWin = window.open(url,'new_page','width=300,height=400')
}
function openAnnounceWindow(url) {
  popupWin = window.open(url,'new_page','width=300,height=400,scrollbars=1')
}

// done hiding -->
</script>
<style type=text/css><!--
a:link    {color:<% =strLinkColor %>;text-decoration:<% =strLinkTextDecoration %>}
a:visited {color:<% =strVisitedLinkColor %>;text-decoration:<% =strVisitedTextDecoration %>}
a:hover   {color:<% =strHoverFontColor %>;text-decoration:<% =strHoverTextDecoration %>}
input.radio {background: <% = strPopUpTableColor %>; color:#000000}
font {  font-size: 9pt; line-height: 13pt; FONT-FAMILY:<% =strDefaultFontFace %>}
td {  font-size: 9pt; line-height: 13pt; FONT-FAMILY:<% =strDefaultFontFace %>}
textarea { BACKGROUND-COLOR: #e8e8e8; BORDER-BOTTOM: 1px double; BORDER-LEFT: 1px double; BORDER-RIGHT: 1px double; BORDER-TOP: 1px double; COLOR: #000000;; font-size: 9pt ;FONT-FAMILY:<% =strDefaultFontFace %>}
input { BACKGROUND-COLOR: #e8e8e8; BORDER-BOTTOM: 1px double; BORDER-LEFT: 1px double; BORDER-RIGHT: 1px double; BORDER-TOP: 1px double; COLOR: #000000;; font-size: 9pt; FONT-FAMILY:<% =strDefaultFontFace %>}

--></style>
<% If intSideMenu = 1 then %>
<!--#include file="popout_script.asp"-->
<% End if
	If intPollMentor = 1 then %>
<!--#include file="inc_pollmentor.asp"-->
<% End If %>
</HEAD>
<%
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
		strOnlineLocation = "回复主题' <a href=""topic.asp?" & strOnlineQueryString & """>" & ChkString(Request.QueryString("Topic_Title"),"title") & "</a> '"
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
		strOnlineLocation = "<a href=""pop_profile.asp?" & strOnlineQueryString & """>正在修改个人资料</a> '"
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
	strSqL = "SELECT count(UserID) AS Guests  "
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
	strSqL = "SELECT count(UserID) AS Members  "
end if
strSql = strSql & "FROM " & strTablePrefix & "ONLINE "
strSql = strSql & " WHERE Right(UserID, 5) <> 'Guest' "

Set rsMembers = my_Conn.Execute(strSql)
Members = rsMembers("Members")
strOnlineMembersCount = rsMembers("Members")

' END WHOS ONLINE SCRIPT
%>
<body background="<%= strPageBGImage %>" topmargin="5" leftmargin="20" bgColor="<% =strPageBGColor %>" text="<% =strDefaultFontColor %>" link="<% =strLinkColor %>" aLink="<% =strActiveLinkColor %>" vLink="<% =strActiveLinkColor %>" <% If intSideMenu = 1 then response.write "onload=init()" end if %> >
<table width="100%" align="center" cellspacing="0" cellpadding="1"  border="0" bgcolor=#000000>
  <tr>
    <td>
      <table width="100%" cellspacing="0" cellpadding="4" border="0">
        <tr>
          <td bgcolor="<% =strForumCellColor %>" valign="middle" align=left>
           <b>>></b> 
<% if strAuthType="nt" then %><b><% =Session(strCookieURL & "username")%>&nbsp;(<% =Session(strCookieURL & "userid") %>)</b>&nbsp;
<%
	else 
		if strAuthType = "db" then %>
<b><% = ChkString(strDBNTUserName, "display") %></b>
<a href="<% =strHomeURL %>"><acronym title="返回主页">主页</acronym></a> | 
<% if strUseExtendedProfile then %>
<a href="pop_profile.asp?mode=Edit"><acronym title="编辑你的个人资料">资料</acronym></a>
<% else %>
<a href="javascript:openWindow3('pop_profile.asp?mode=Edit')"><acronym title="编辑你的个人资料">资料</acronym></a>
<%
			end if 
			if strAutoLogon <> 1 then %> | <a href="policy.asp"><acronym title="注册了才能发表文章哦！">注册</acronym></a><% end if %>
        |
        <a href="active.asp"><acronym title="看看有什么新文章吧！">最新</acronym></a>
        |
        <a href="active_users.asp"><acronym title="看看有谁在线..">在线</acronym></a>
        |
        <a href="topic_stats.asp"><acronym title="查看主题相关统计数字">统计</acronym></a>
        |
        <a href="members.asp"><acronym title="论坛现有会员">会员</acronym></a>
        |
        <a href="search.asp"><acronym title="按关键字、日期、作者来搜寻">搜索</acronym></a>
        |
        <a href="bookmark.asp"><acronym title="查看收藏的贴子">收藏</acronym></a>
        |
        <a href="faq.asp"><acronym title="常见问题的解答">帮助</acronym></a>
        |
        <a href="cancel_replies.asp"><acronym title="取消提醒电子邮件">设置</acronym></a>
        |
        <a href="events.asp"><acronym title="查看或新增日记">日记</acronym></a></font>
<% if (lcase(strNoCookies) = "1") then %> | <a href="admin_home.asp">管理选项</a>
<% end if %>
<% if (mlev = 4) or (lcase(strNoCookies) = "1") then %>
 | <a href="admin_home.asp">管理</a>
<% end if %>
          </td>
<td bgcolor="<% =strForumCellColor %>" valign="middle" align=right><b><% =strForumTitle %></b></td>
            </tr>
        </table>
        </td>
    </tr>
</table>
<% If intSideMenu = 1 then %>
<!--#include file="popout_menu.asp"-->
<% End If %>
<a name="top"></a>

<font face="<% =strDefaultFontFace %>">
<table align="center" border="0" cellPadding="0" cellSpacing="0" width="100%">
  <tr>
    <td valign="top" width="30%"><a href="default.asp"><img alt="<% =strForumTitle %>" border="0" src="<%=strTitleImage%>"></a>
<% '################ Poll Mentor
	if intPollMentor = 1 then
			Dim aTitle
			Dim aQuestion
			Dim aAnswers(8)
			Dim aCount(8)
			'Get active one...
			PollMentor_GetPollInfo -1, aTitle, aQuestion, aAnswers, aCount %>
			<div align="center" class="PollHeader" id="PollHeader" onmouseup="javascript:openpollWindow('pollmentor.asp')">
			<TABLE  BORDER=0 bgcolor="<% =strTableBorderColor %>" cellspacing="1" cellpadding="3" width="100%">
			<tr onmouseup="javascript:openpollWindow('pollmentor.asp')">
			      <TD onmouseup="javascript:openpollWindow('pollmentor.asp')" colspan="2" bgcolor="<% =strCategoryCellColor %>" ALIGN="LEFT"><b><font face="<% =strDefaultFontFace %>" color="<% =strHeadFontColor %>" size="<% =strDefaultFontSize %>"><%= aTitle %><%= aQuestion %>&nbsp;&nbsp; <a href="Javasript:;" onclick="javascript:openpollWindow('pollmentor.asp')" >按这里投票</a></font></b></TD>
			</tr>
			</table>
			</div>
<%	end if
'################ Poll Mentor %>
<br>
<%'########################################################################
  if InstrRev(Request.ServerVariables("Path_Info"), "view_announcements.asp") = 0 then
    if InstrRev(Request.ServerVariables("Path_Info"), "admin_") = 0 then
	'## Forum_SQL - Build SQL to get forums via category
	strSql = "SELECT " & strTablePrefix & "ANNOUNCE.A_ID "
	strSql = strSql & "FROM " & strTablePrefix & "ANNOUNCE " 

	set rschkAnnounce = my_Conn.Execute (strSql)

	if rschkAnnounce.eof or rschkAnnounce.bof then '## No Announcements found in DB
		Response.Write ""
	Else
%><TABLE  BORDER=0 bgcolor="<% =strPageBGColor %>" cellspacing="1" cellpadding="3" width="100%">
	<tr align="center">
          <td>
            <!--#INCLUDE FILE="inc_announce.asp" -->    
	 </td>
       </tr>
	  </table>
<%	End If
	set rschkAnnounce = nothing
    end if
  end if
'#########################################################################################%>


    </td>
    <td align="center" valign="top" width="70%">
    <table border="0" cellPadding="2" cellSpacing="0">
      <tr>
        <td align="center">
<%
select case Request.Form("Method_Type")

	case "login"
%>
    </td></tr></table></td></tr></table>
<%		if strLoginStatus = 0 then %>
<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">你的用户名或密码错误</font></p>

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">请重新输入或注册成为新会员</font></p>
<%		else %>
<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">登陆成功！！</font></p>

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">感谢你的参与！</font></p>
<%		end if %>
<meta http-equiv="Refresh" content="2; URL=<% =Request.ServerVariables("HTTP_REFERER") %>">

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="<% =Request.ServerVariables("HTTP_REFERER") %>">返回论坛</font></a></p>

<!--#INCLUDE FILE="inc_footer.asp" -->
<%		Response.End %>
<%	case "logout" %>
    </table>
    </td>
  </tr>
</table>

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">成功退出！！</font></p>

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">有时间记得多来走走哦！</font></p>

<meta http-equiv="Refresh" content="2; URL=<% =Request.ServerVariables("HTTP_REFERER") %>">

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="<% =Request.ServerVariables("HTTP_REFERER") %>">返回论坛</font></a></p>

<!--#INCLUDE FILE="inc_footer.asp" -->
<%
		Response.End
		end select

if (ChkUser2((strDBNTUserName), (Request.Cookies(strUniqueID & "User")("Pword"))) = 0) then %>
      <form action="<% =Request.ServerVariables("URL") %>" method="post" id=form1 name=form1>
      <INPUT type="hidden" name="Method_Type" value="login">
      <tr>
        <td align="center">
        <table border=0>
          <tr>
<%	if (strAuthType = "db") then %>
            <td><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">用户名：</font> 
            <INPUT type="text" name="Name" size="10" value=""></td>
            <td><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">密码：</font> 
            <INPUT type="password" name="Password" size="10" value=""></td>
            <td valign="bottom">
<%
		if strGfxButtons <> "0" then %>
            <INPUT src=<%= strImageURL %>button_login.gif type="image" value="登陆" id=submit1 name=submit1 border=0 hspace=4>
<%	else
			if strGfxButtons = "0" then %>
            <INPUT type="submit" value="登陆" id=submit1 name=submit1>
<%			end if
		end if %>
            </td>
	</tr>
	<tr>
	<td colspan="3" align="left"><font face="<% =strDefaultFontFace %>" size="<%=strFooterFontSize %>">
	<INPUT type="checkbox" name="SavePassWord" value="true" CHECKED>记住密码</font>
<%else
		if (strAuthType = "nt") then %>
 | <font face="<% =strDefaultFontFace %>" size="1" color="red"><a href="policy.asp">注册</a>了才能发表</font>
<%		end if
	end if %>
<%
	if (lcase(strEmail) = "1") then %>
 | <font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>"><a href="JavaScript:openWindow('pop_pword.asp')">忘记了<% if strAuthType = "nt" then %>管理<% end if %>密码？</a></font></td>
         </tr>
<%
	end if 
%>
       </table>
       </td>
     </tr>
     </form>
<%
else 
%>
     <form action="<% =Request.ServerVariables("URL") %>" method="post" id=form2 name=form2>
     <INPUT type="hidden" name="Method_Type" value="logout">
     <tr>
       <td align="center">
        <table border=0>
          <tr>
            <td align="center">

            </td>
            <td>
<%
			if strGfxButtons <> "0" then %>
            <INPUT src=<%= strImageURL %>button_logout.gif type="image" value="Logout" id=submit1 name=submit1 border=0 hspace=4> 
<%
			else
				if strGfxButtons = "0" then %>
            <INPUT type="submit" value="Logout" id=submit1 name=submit1> 
<%				end if 
			end if 
		end if 
	end if %>
            </td>
          </tr>
        </table>
        </td>
     </tr>

     </form>
	 <tr><td align="right"></td></tr>
<% end if %>

</table>
<% If intPMessages = 1 then %>
<br><!--#include file="privatemess.asp"-->
<% End If %>
   </td>
 </tr>
</table>
