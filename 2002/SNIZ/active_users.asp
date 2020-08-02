<!--#INCLUDE FILE="config.asp" -->
<!--#INCLUDE FILE="inc_functions.asp" -->
<!--#INCLUDE FILE="inc_top.asp" -->
<%
'## Do Cookie stuffs with reload
nRefreshTime = Request.Cookies(strTempCookieType & "Reload")

if Request.form("cookie") = "1" then
        if strSetCookieToForum = 1 then
      Response.Cookies(strTempCookieType & "Reload").Path = strTempCookieType
	end if
	Response.Cookies(strTempCookieType & "Reload") = Request.Form("RefreshTime")
	Response.Cookies(strTempCookieType & "Reload").expires = strForumTimeAdjust + 365
	nRefreshTime = Request.Form("RefreshTime")
end if

if nRefreshTime = "" then
	nRefreshTime = 0
end if
ActiveSince = Request.Cookies(strTempCookieType & "ActiveSince")
'## Do Cookie stuffs with show last date
if Request.form("cookie") = "2" then
	ActiveSince = Request.Form("ShowSinceDateTime")
    if strSetCookieToForum = 1 then
      Response.Cookies(strTempCookieType & "ActiveSince").Path = strTempCookieType
	end if
	Response.Cookies(strTempCookieType & "ActiveSince") = ActiveSince
end if


mypage = request("whichpage")

	If  mypage = "" then
	   mypage = 1
	end if

	mypagesize = request.cookies("paging")("pagesize")

	If  mypagesize = "" then
	   mypagesize = 15
	end if
%>
<script language="JavaScript">
<!--
function autoReload()
{
	document.ReloadFrm.submit()
}
//-->
</script>
<script language="JavaScript">
<!--
function SetLastDate()
{
	document.LastDateFrm.submit()
}
//-->
</script>
<script language="JavaScript">
<!--
function jumpTo(s) {if (s.selectedIndex != 0) top.location.href = s.options[s.selectedIndex].value;return 1;}
// -->
</script>
<table width="100%" border=0>
  <tr>
    <td ><img src="<%=strImageURL %>icon_folder_open_topic.gif" border="0">&nbsp;论坛在线会员列表</font><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">（最后更新时间：<% =chkDate(DateToStr(strForumTimeAdjust)) %> <% =chkTime(DateToStr(strForumTimeAdjust)) %>）</font>
    </td>
    <td align="right">
    <form name="ReloadFrm" action="active_users.asp" method="post"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    <select name="RefreshTime" size="1" onchange="autoReload();">
        <option value="0"  <% if nRefreshTime = "0" then Response.Write(" SELECTED")%>>不要刷新</option>
        <option value="1"  <% if nRefreshTime = "1" then Response.Write(" SELECTED")%>>每分钟刷新一次</option>
        <option value="5"  <% if nRefreshTime = "5" then Response.Write(" SELECTED")%>>每隔五分钟刷新一次</option>
        <option value="10" <% if nRefreshTime = "10" then Response.Write(" SELECTED")%>>每隔十分钟刷新一次</option>
        <option value="15" <% if nRefreshTime = "15" then Response.Write(" SELECTED")%>>每隔十五分钟刷新一次</option>
        <option value="30" <% if nRefreshTime = "30" then Response.Write(" SELECTED")%>>每隔半个小时刷新一次</option>
    </select>
    <input type="hidden" name="Cookie" value="1">
    </font>
    </form>
    </td>
  </tr>
</table>
<CENTER>
<SCRIPT>
<!--
if (document.ReloadFrm.RefreshTime.options[document.ReloadFrm.RefreshTime.selectedIndex].value > 0) {
	reloadTime = 60000 * document.ReloadFrm.RefreshTime.options[document.ReloadFrm.RefreshTime.selectedIndex].value
	self.setInterval('autoReload()', 60000 * document.ReloadFrm.RefreshTime.options[document.ReloadFrm.RefreshTime.selectedIndex].value)
}
//-->
</SCRIPT>
<P><B><FONT SIZE="4" FACE="宋体, Arial" COLOR="<% =strDefaultFontColor %>">论坛在线会员列表</FONT></B><P>
<table bgcolor="<% =strTableBorderColor %>" cellpadding=2 border=0 cellspacing=1 width="100%">
  <tr>
    <td bgcolor="<% =strHeadCellColor %>" align=center valign=top nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>"><B>会员名</B></FONT></td>
<%	if (mlev = 4 or mlev = 3) then %>
    <td bgcolor="<% =strHeadCellColor %>" align=center valign=top nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>"><B>IP 地址</B></FONT></td>
<%	end if %>
	<td bgcolor="<% =strHeadCellColor %>" align=center valign=top nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>"><B>状态</B></FONT></td>
    <td bgcolor="<% =strHeadCellColor %>" align=center valign=top nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>"><B>登陆时间</B></FONT></td>
    <td bgcolor="<% =strHeadCellColor %>" align=center valign=top nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>"><B>最近活动时间</B></FONT></td>
    <td bgcolor="<% =strHeadCellColor %>" align=center valign=top nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>"><B>停留时间</B></FONT></td>
</tr>
<%
	set rs = Server.CreateObject("ADODB.Recordset")
	'## Forum_SQL
	strSql ="SELECT " & strTablePrefix & "ONLINE.UserID, " & strTablePrefix & "ONLINE.UserIP, " & strTablePrefix & "ONLINE.M_BROWSE, " & strTablePrefix & "ONLINE.DateCreated, " & strTablePrefix & "ONLINE.LastChecked, " & strTablePrefix & "ONLINE.CheckedIn "
	strSql = strSql & " FROM " & strMemberTablePrefix & "ONLINE "
	strSql = strSql & " ORDER BY " & strTablePrefix & "ONLINE.DateCreated, " & strTablePrefix & "ONLINE.CheckedIn DESC"

	rs.cachesize = 20
	rs.open  strSql, my_Conn, 3

	i = 0

	If rs.EOF or rs.BOF then  '## No categories found in DB
		Response.Write ""
	Else

		rs.movefirst
		num = 0
		rs.pagesize = mypagesize
		maxpages = cint(rs.pagecount)
		maxrecs = cint(rs.pagesize)
		rs.absolutepage = mypage
		howmanyrecs = 0
		rec = 1
		do until rs.EOF or (rec = mypagesize+1)
			if strI = 0 then
				CColor = strAltForumCellColor
			else
				CColor = strForumCellColor
			end if

  			strTheUserID = rs("UserID")
  			strTheUserID = OnlineSQLdecode(strTheUserID)

  			if Right(rs("UserID"), 5) <> "Guest" then
	strSql = "SELECT "   & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME,  " & strTablePrefix & "ONLINE.UserID "
	strSql = strSql & " FROM " & strTablePrefix & "MEMBERS, " & strTablePrefix & "ONLINE "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & rs("UserID") & "' "
	set rsMember =  my_Conn.Execute (strSql)
end if
if Right(rs("UserID"), 5) = "Guest" then
num = num + 1
end if

strRSCheckedIn = rs("CheckedIn")
strOnlineLastDateChecked = rs("LastChecked")
strOnlineDateCheckedIn = StrToDate(strRSCheckedIn)
strOnlineLastDateChecked = StrToDate(strOnlineLastDateChecked)

strOnlineTotalTime = DateDiff("n",strOnlineDateCheckedIn,strOnlineLastDateChecked)

If strOnlineTotalTime > 60 then
' they must have been online for like an hour or so.
strOnlineHours = 0
do until strOnlineTotalTime < 60
strOnlineTotalTime = (strOnlineTotalTime - 60)
strOnlineHours = strOnlineHours + 1
loop
strOnlineTotalTime = strOnlineHours & " 小时 " & strOnlineTotalTime & " 分钟"
Else
strOnlineTotalTime = strOnlineTotalTime & " 分钟"
End If

%>
  <tr bgcolor="<% =CColor %>">
<%  if Right(rs("UserID"), 5) = "Guest" then %>
    <td valign="middle" align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strDefaultFontColor %>">游客 #<% =num %></font></td>
    <% else %>
    <td valign="middle" align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strDefaultFontColor %>">
<%          if strUseExtendedProfile then
				Response.Write("<a href=""pop_profile.asp?mode=display&id="& rsMember("MEMBER_ID") & """>")
			else
				Response.Write("<a href=""JavaScript:openWindow2('pop_profile.asp?mode=display&id=" & rsMember("MEMBER_ID") & "')"">")
			end if
			Response.Write(rs("UserID") & "</a></font>")
%>
    <% end if %>
<%	if (mlev = 4 or mlev = 3) then %>
	<td valign="middle" align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strDefaultFontColor %>"><% =rs("UserIP")%></A></font></td>
<%	end if %>
	<td valign="middle" align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strDefaultFontColor %>"><% =rs("M_BROWSE")%></A></font></td>
    <td align="center" valign="middle" nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strDefaultFontColor %>"><% =chkTime(strRSCheckedIn) %></FONT></td>
    <td align="center" valign="middle" nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strDefaultFontColor %>"><% =chkTime(DateToStr(strOnlineLastDateChecked)) %></FONT></td>
    <td align="center" valign="middle" width="100" nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strDefaultFontColor %>"><%=strOnlineTotalTime%></FONT></td>
  </tr>
<%
		    rs.MoveNext
		    strI = strI + 1
		    if strI = 2 then
				strI = 0
			end if
		    rec = rec + 1
		loop
	end if
 %>

</table>
<table width="90%">
<tr>
     <td>&nbsp;</td>
     <td align=right width=100>
     <% if maxpages > 1 then %>
	<table width="95%" border=0 align="right">
  	<tr>
    	<td valign="top"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>页数：</b></font></td>
    	<td valign="top"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% Call Paging() %></font></td>
  	</tr>
	</table>
<% else %>

    &nbsp;
<% end if %>
</td>
</tr>
</table>
<p>
<center>

<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>-1" color="<% =strCategoryFontColor %>">
<a href="default.asp">返回论坛</A>
</FONT>
<p>
<br>
<!--#INCLUDE FILE="inc_footer.asp" -->

<%
sub Paging()
	if maxpages > 1 then
		if Request.QueryString("whichpage") = "" then
			pge = 1
		else
			pge = Request.QueryString("whichpage")
		end if
		scriptname = request.servervariables("script_name")
		Response.Write("<table border=0 width=100% cellspacing=0 cellpadding=1 align=top><tr>")
		for counter = 1 to maxpages
			if counter <> cint(pge) then
				ref = "<td align=right bgcolor=" & strPageBGColor & "><font face=" & strDefaultFontFace & " size=" & strDefaultFontSize & ">" & "&nbsp;" & widenum(counter) & "<a href='" & scriptname
				ref = ref & "?whichpage=" & counter
				ref = ref & "&pagesize=" & request.cookies("paging")("pagesize")
				if top = "1" then
					ref = ref & ">"
					ref = ref & "</font><b><font face='" & strDefaultFontFace & "' "
					ref = ref & "color='" & strHeadFontColor & "'"
					ref = ref & ">" & counter & "</font></b></a></td>"
					Response.Write ref
				else
					ref = ref & "'>" & counter & "</font></a></td>"
					Response.Write ref
				end if
			else
				Response.Write("<td align=right bgcolor=" & strPageBGColor & "><font face=" & strDefaultFontFace & " size=" & strDefaultFontSize & ">" & "&nbsp;" & widenum(counter) & "<b>" & counter & "</b></font></td>")
			end if
			if counter mod 15 = 0 then
				Response.Write("</tr><tr>")
			end if
		next
		Response.Write("</tr></table>")
	end if
	top = "0"
end sub
 %>