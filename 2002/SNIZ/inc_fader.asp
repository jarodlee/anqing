<tr>
   <td bgcolor="<% =strTableBorderColor %>">
       <table border=0 width="100%" cellspacing=1 cellpadding=4>
	<tr>
	  <td bgcolor="<% =strCategoryCellColor %>" colspan="1"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" valign="top">论坛滚动新闻</font></td>
        </tr>
	<tr>
	  <td bgcolor="<% =strForumCellColor %>" align=center valign="center" colspan="1">
	    <applet code="Fade.class" width="500" height="30">
	      <param name="bgcolor" value="ffffff">
	      <param name="txtcolor" value="000000">
	      <param name="changefactor" value="2">
	      <param name="text1" value="欢迎来到<% =strForumTitle %>">
	      <param name="url1" value="">
	      <param name="font1" value="宋体,PLAIN,12">
	      <param name="text2" value="今日新闻">
	      <param name="url2" value="">
	      <param name="font2" value="宋体,PLAIN,12">
<%
	'## Forum_SQL
	strSql = "SELECT " & strTablePrefix & "ANNOUNCE.A_ID"
	strSql = strSql & ", " & strTablePrefix & "ANNOUNCE.A_AUTHOR"
	strSql = strSql & ", " & strTablePrefix & "ANNOUNCE.A_SUBJECT"
	strSql = strSql & ", " & strTablePrefix & "ANNOUNCE.A_MESSAGE"
	strSql = strSql & ", " & strTablePrefix & "ANNOUNCE.A_START_DATE"
	strSql = strSql & ", " & strTablePrefix & "ANNOUNCE.A_END_DATE"
	strSql = strSql & " FROM " & strTablePrefix & "ANNOUNCE "
	if mLev < 4 then
		strSql = strSql & " WHERE " & strTablePrefix & "ANNOUNCE.A_START_DATE <= " & "'" & DatetoStr(Now()) & "'"
		strSql = strSql & " AND " & strTablePrefix & "ANNOUNCE.A_END_DATE > " & "'" & DatetoStr(Now()) & "'"
        end if
	strSql = strSql & " ORDER BY " & strTablePrefix & "ANNOUNCE.A_START_DATE DESC"
	strSql = strSql & ", " & strTablePrefix & "ANNOUNCE.A_ID DESC;"

	set rsAnnounce = Server.CreateObject("ADODB.Recordset")
	rsAnnounce.cachesize = 20
	rsAnnounce.open  strSql, my_Conn, 3

	if not(rsAnnounce.EOF or rsAnnounce.BOF) then  '## Replies found in DB
		rsAnnounce.movefirst
		rsAnnounce.pagesize = strPageSize
		maxpages = cint(rsAnnounce.pagecount)
	end if
	if rsAnnounce.EOF or rsAnnounce.BOF then  '## No replies found in DB
		Response.Write ""
	else
		'rsAnnounce.movefirst
		intI = 3 

		do until rsAnnounce.EOF '**
			AnnounceSubject = rsAnnounce("A_SUBJECT")
			AnnounceLink = strForumURL & "view_announcements.asp" %>
	      <param name="text<% =intI %>" value="<% =AnnounceSubject %>">
	      <param name="url<% =intI %>" value="<% =AnnounceLink %>">
	      <param name="font<% =intI %>" value="宋体,PLAIN,12">
<%
		    rsAnnounce.MoveNext
		    intI  = intI + 1
		loop
	end if
%>
 	    </applet></td>
	</tr>
     </table>
   </td>
</tr>