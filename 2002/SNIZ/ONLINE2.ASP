
<script language="JavaScript">
<!-- hide from JavaScript-challenged browsers
function pop_online() { var MainWindow = window.open ("pop_online.asp", "","toolbar=no,location=no,menubar=no,scrollbars=yes,width=250,height=500,resizeable=yes,status=no");
}
// done hiding -->
</script>
<% 
if strShowStatistics <> "1" then %>
<tr>
<td bgcolor="<% =strCategoryCellColor %>" colspan="<%=sGetColspan(7, 6)%>">
	  <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" size="+1">
<b>论坛在线会员</b></font>
</td>
</tr>
<% End If %>
 <TR>
	<td align="center" bgcolor="<% =strForumCellColor %>" valign="middle">
		<a HREF="javascript:pop_online()"><img src="<%=strImageURL %>online.gif" border="0"></a>
	</td>
<TD valign=top bgcolor="<% =strForumCellColor %>" colspan="<%=sGetColspan(6, 5)%>">
<%

 if Request.Cookies(strUniqueID & "User")("Name") = "" Then %>
	<font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>" color="<% =strActiveLinkColor %>">
    <a HREF="active_users.asp"><b>论坛在线会员</b></a><br></font>
<%	else %>
	<font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>" color="<% =strActiveLinkColor %>">
    <a HREF="active_users.asp"><b>论坛在线会员</b></a><br>

<%		strTemp = "有 "
		strGuest = "位游客"



	set rs = Server.CreateObject("ADODB.Recordset")
	'## Forum_SQL
	strSql ="SELECT " & strTablePrefix & "ONLINE.UserID, " & strTablePrefix & "ONLINE.M_BROWSE, " & strTablePrefix & "ONLINE.DateCreated "
	strSql = strSql & " FROM " & strMemberTablePrefix & "ONLINE "
	strSql = strSql & " ORDER BY " & strTablePrefix & "ONLINE.DateCreated DESC"
 
	rs.cachesize = 100
	rs.open  strSql, my_Conn, 3

	i = 0 

	If rs.EOF or rs.BOF then  '## No categories found in DB
		Response.Write ""
	Else

		rs.movefirst
		do until rs.EOF 
  if Right(rs("UserID"), 5) <> "Guest" then 
	strSql = "SELECT "   & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME,  " & strTablePrefix & "ONLINE.UserID "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS, " & strTablePrefix & "ONLINE "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & rs("UserID") & "' AND " & strTablePrefix & "ONLINE.UserID = M_NAME" 
	set rsMember =  my_Conn.Execute (strSql)
	end if


  if Right(rs("UserID"), 5) <> "Guest" then %>
    <font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>" color="<% =strForumLinkFontColor %>"><% = chkString(rs("UserID"),"name")%>, </font>
<%			
    end if
		    rs.MoveNext
		    rec = rec + 1
		loop
			Response.Write ( " 和 " & strOnlineGuestsCount & " " & strGuest )
	end if
end if 
%>
	
	</font>
</TD>
</TR>
