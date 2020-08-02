<!--#INCLUDE FILE="../config.asp" -->
<font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
<%
	set my_Conn = Server.CreateObject("ADODB.Connection")
	my_Conn.Open strConnString
	on error resume next

	strSql = "DROP TABLE " & strTablePrefix & "ONLINE "

	my_Conn.Execute strSql
	response.write(err.number & " | " & err.description & "<br>")
	strSql = "DELETE FROM " & strTablePrefix & "MODS "

	my_Conn.Execute strSql
	response.write(err.number & " | " & err.description & "<br>")


	my_Conn.Close
	set my_Conn = nothing
	Response.End
%>
</font>