<!--#INCLUDE FILE="config.asp" -->
<!--#INCLUDE FILE="inc_functions.asp" -->
<!--#INCLUDE FILE="inc_top.asp" -->
<table align=left border=0><tr><td width="33%" align="left" nowrap>
	<tr><td><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    <a href="default.asp"><img src="<%=strImageURL %>icon_folder_open.gif" alt="所有论坛" height=15 width=15 border="0">&nbsp;返回论坛首页</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" height=15 width=15 border="0">&nbsp;主题统计</font>
    </font></td>
	</tr>
	<tr><td align=center>
<%
sort_order=request("sort_order")
if sort_order="" then
sort_order="T_VIEW_COUNT"
end if



SQL2 = "SELECT T_SUBJECT, T_VIEW_COUNT, T_REPLIES, TOPIC_ID, T_DATE, T_AUTHOR, T_LAST_POST_AUTHOR FROM FORUM_TOPICS where "&sort_order&">0 order by "&sort_order &" DESC"
	set onn = Server.CreateObject("ADODB.Connection")
	onn.Open strConnString 
	set rs2 = onn.execute(SQL2)

%>
<table border="0" cellpadding="2" cellspacing="1" width="100%" bgcolor="<% =strTableBorderColor %>">
  <tr>
    <td width="38%" bgcolor="<%= strCategoryCellColor %>" align="Center"><font color="<%= strCategoryFontColor %>" size="<%= strDefaultFontSize %>"><strong>主题</strong></font></td>
    <td width="10%" bgcolor="<%= strCategoryCellColor %>" align="Center"><font color="<%= strCategoryFontColor %>" size="<%= strDefaultFontSize %>"><strong>作者</strong></font></td>
    <td width="15%" bgcolor="<%= strCategoryCellColor %>" align="left"><font color="<%= strCategoryFontColor %>" size="<%= strDefaultFontSize %>"><strong>日期</strong></font></td>
    <td width="5%" bgcolor="<%= strCategoryCellColor %>" align="left"><font color="<%= strCategoryFontColor %>" size="<%= strDefaultFontSize %>"><strong><a href="topic_stats.asp?sort_order=T_VIEW_COUNT">人气</a></strong></td>
    <td width="5%" bgcolor="<%= strCategoryCellColor %>" align="left"><font color="<%= strCategoryFontColor %>" size="<%= strDefaultFontSize %>"><strong><a href="topic_stats.asp?sort_order=T_REPLIES">回复</a></strong></td>
    <td width="10%" bgcolor="<%= strCategoryCellColor %>" align="Center"><font color="<%= strCategoryFontColor %>" size="<%= strDefaultFontSize %>"><strong>最后回复</strong></font></td>
 </tr>
 <%
 	x=1
	do until rs2.eof
	
	if x mod 2 then
	tdcolor= strForumCellColor
	else
	tdcolor= strAltForumcellColor
	end if
	x=x+1
	%>
	<td width="38%" bgcolor="<%=tdcolor%>" valign=top><font size="<%= strFooterFontSize+1 %>">
	<a href="link.asp?TOPIC_ID=<%=rs2("TOPIC_ID")%>">
	<%=rs2("T_SUBJECT")%></a></td>
	<td width="10%" bgcolor="<%=tdcolor%>" valign=top><font size="<%= strFooterFontSize+1 %>">
	<a href="member_profile.asp?mode=display&id=<%=rs2("T_AUTHOR")%>">
	<%=getMemberName(rs2("T_AUTHOR"))%></a></td>
	<td width="12%" bgcolor="<%=tdcolor%>" valign=top><font size="<%= strFooterFontSize+1 %>">
<font size=1>	<%=chkdate(rs2("T_DATE"))%>, <%=chktime(rs2("T_DATE"))%></td>
	<td width="5%" bgcolor="<%=tdcolor%>" valign=top><font size="<%= strFooterFontSize+1 %>">
	<%=rs2("T_VIEW_COUNT")%></td>
	<td width="5%" bgcolor="<%=tdcolor%>" valign=top><font size="<%= strFooterFontSize+1 %>">
	<%=rs2("T_REPLIES")%></td>
	<td width="5%" bgcolor="<%=tdcolor%>" valign=top><font size="<%= strFooterFontSize+1 %>">
	<%
	if rs2("T_LAST_POST_AUTHOR")<>0 then
	response.write "<a href=""member_profile.asp?mode=display&id="&rs2("T_LAST_POST_AUTHOR")&""">"
	response.write getmemberName(rs2("T_LAST_POST_AUTHOR"))&"</a> " 
	else
	response.write "&nbsp;"
	end if
	%></td><tr>
	<%
	rs2.movenext
	loop
	
rs2.close
set rs2=nothing
onn.close
%>
</table>
</td>
</tr>
<tr>
<br>
<!--#INCLUDE FILE="inc_footer.asp" -->
