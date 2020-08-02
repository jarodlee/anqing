<!--#INCLUDE FILE="../config.asp" -->
<font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
<%
	set my_Conn = Server.CreateObject("ADODB.Connection")
	my_Conn.Open strConnString
	on error resume next
	
	strSQL = "DELETE FROM " & strTablePrefix & "MODS WHERE M_NAME = 'Attachment' OR M_NAME='HModEnable' OR M_NAME='NewMemPM'"
	my_conn.Execute strSQL
	
	response.write(err.number & " | " & err.description & "<br>")
	strSql = "INSERT INTO " & strTablePrefix & "MODS (M_NAME, M_CODE, M_VALUE) values ('"
	strSql = strSql & "Attachment','faMaxSize','512')"
	my_Conn.Execute strSql
	response.write(err.number & " | " & err.description & "<br>")
	
	strSql = "INSERT INTO " & strTablePrefix & "MODS (M_NAME, M_CODE, M_VALUE) values ('"
	strSql = strSql & "Attachment','faExtensions','.zip;.mdb')"
	my_Conn.Execute strSql
	response.write(err.number & " | " & err.description & "<br>")

	strSql = "INSERT INTO " & strTablePrefix & "MODS (M_NAME, M_CODE, M_VALUE) values "
	strSql = strSql & "('NewMemPM','Subject','Enter your subject line here')"
	my_Conn.Execute strSql
	strSql = "INSERT INTO " & strTablePrefix & "MODS (M_NAME, M_CODE, M_VALUE) values "
	strSql = strSql & "('NewMemPM','Message','Enter your message here')"
	my_Conn.Execute strSql
	strSql = "INSERT INTO " & strTablePrefix & "MODS (M_NAME, M_CODE, M_VALUE) values "
	strSql = strSql & "('NewMemPM','OnOff','0')"
	my_Conn.Execute strSql

	strSql = "INSERT INTO " & strTablePrefix & "MODS (M_NAME, M_CODE, M_VALUE) values "
	strSql = strSql & "('HModEnable','Attachment','0')"
	my_Conn.Execute strSql
	strSql = "INSERT INTO " & strTablePrefix & "MODS (M_NAME, M_CODE, M_VALUE) values "
	strSql = strSql & "('HModEnable','Pmessages','1')"
	my_Conn.Execute strSql
	strSql = "INSERT INTO " & strTablePrefix & "MODS (M_NAME, M_CODE, M_VALUE) values "
	strSql = strSql & "('HModEnable','UserFields','1')"
	my_Conn.Execute strSql
	strSql = "INSERT INTO " & strTablePrefix & "MODS (M_NAME, M_CODE, M_VALUE) values "
	strSql = strSql & "('HModEnable','PollMentor','0')"
	my_Conn.Execute strSql
	strSql = "INSERT INTO " & strTablePrefix & "MODS (M_NAME, M_CODE, M_VALUE) values "
	strSql = strSql & "('HModEnable','SideMenu','0')"
	my_Conn.Execute strSql
	strSql = "INSERT INTO " & strTablePrefix & "MODS (M_NAME, M_CODE, M_VALUE) values "
	strSql = strSql & "('HModEnable','NewMemberPM','1')"
	my_Conn.Execute strSql
	strSql = "INSERT INTO " & strTablePrefix & "MODS (M_NAME, M_CODE, M_VALUE) values "
	strSql = strSql & "('HModEnable','imageURLPath','/')"
	my_Conn.Execute strSql

	response.write(err.number & " | " & err.description & "<br>")

	response.write "<h1>Database set-up finished..</h1>"

	my_Conn.Close
	set my_Conn = nothing
	Response.End
%>

If the last ALTER TABLE fails, you will need to set a default value of 1 for M_ALLOWUPLOADS and M_ALLOWDOWNLOADS

</font>