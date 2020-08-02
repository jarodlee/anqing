<!--#include file="../config.asp"-->
<!--#include file="../inc_top_short.asp"-->

<%

	Set rs = Server.CreateObject("ADODB.Recordset")
	if request("emin") = "New" then
		set rs.ActiveConnection = my_Conn
		rs.CursorType = 3
		my_Conn.Execute = "INSERT INTO FORUM_USERFIELDS (USR_LABEL,USR_FIELDTYPE,USR_SHORTNAME) Values ('NewLabel','S','User_')"
		response.redirect "../admin_user_fields.asp"
	end if
%>

<%
	if request("emin") = "Del" then
		sfldValue = request("USR_FIELD_ID")

		if isNumeric(sfldValue) then
			sWhere = "USR_FIELD_ID" & "=" & sfldValue
		else
			sWhere = "USR_FIELD_ID" & "='" & sfldValue & "'"
		end if
		my_Conn.execute "delete from  FORUM_USERFIELDS  WHERE " & sWhere
		my_Conn.execute "delete from  FORUM_MEMBERFIELDS  WHERE " & sWhere
		
		response.redirect "../admin_user_fields.asp?uTable=" & request("uTable") & "&iPage="  & request("iPage")
		response.end
	end if
%>

<br>
<br>
<div align="center"><% if request("preview") = "preDel" then  %>

			<p>Are you sure you want to delete the record?</p>
			<a href="pop_ufield_table.asp?<%=request.querystring%>&emin=Del" >Yes</a>&nbsp;
<% Elseif request("preview") = "preNew" then  %>
			<p>Are you sure you want to Add a new record?</p>
			<a href="pop_ufield_table.asp?<%=request.querystring%>&emin=New" >Yes</a>&nbsp;			
<% End If %>

			<a href="../admin_user_fields.asp?uTable=<%=request("uTable")%>&iPage=1">No</a>
<br>
</div></font>
<br>
<br>
<br>
<br>
<br>
<div align="center">
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" >
<a href="../admin_user_fields.asp?uTable=FORUM_USERFIELDS">Back to Table Editor</a>
</font>
</div>
</body>
</html>