<!--#include file="../config.asp"-->
<!--#include file="../inc_top_short.asp"-->

<%

	Set rs = Server.CreateObject("ADODB.Recordset")

	set rs.ActiveConnection = my_conn
	rs.CursorType = 3
	

	sfldName = ""
	sfldValue = request.querystring(2)

	
	sSQL = "SELECT * FROM FORUM_USERFIELDS WHERE USR_FIELD_ID=" & sfldValue
	rs.Open sSQL, , , &H0002
	if request("update") <> "" then
		on error resume next
		
		for each fld in rs.fields
			rs(fld.name) = request(fld.name)
		next
		rs.update
		
		response.redirect "../admin_user_fields.asp?uTable=" & request("uTable") & "&iPage=" & request("iPage")
	end if
	
%>
<br><br>
		<table align="center" border=0 cellspacing=1 cellpadding=2 bgcolor = "<% =strHeadCellColor %>" width=300>		
				<tr>
					
					<td bgcolor="<% =strHeadCellColor %>"  colspan=2><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">UserFields&nbsp;&nbsp;
					</td>
				
				</tr>
		<form action="pop_ufield_record.asp?<%=request.querystring%>&update=1" method="post" name="editform">
		<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" >
		<%
			lRecs = rs.RecordCount
			lFields = rs.Fields.count

			for each fld in rs.fields
				if fld.name <> "USR_FIELD_ID" then
					response.write "<tr>"
					response.write "<td bgcolor=""" & strPopUpTableColor  & """ width=50><font face=" & strDefaultFontFace  & " size=" & strDefaultFontSize  & " color=" & strDefaultFontColor & ">" & fld.name & "</font></td>"
					if fld.name <> "USR_FIELDTYPE" then
					response.write "<td bgcolor=""" & strPopUpTableColor  & """><input size=24 type=""text"" name=""" & fld.name &  """ value=""" & rs(fld.name) &  """></td>"
					else %>
					<td bgcolor="<%= strPopUpTableColor %>">
					<select name="<%= fld.name %>" size="1">
					<% If rs(fld.name) = "S" then %>
					<option value="S" SELECTED>Simple Input</option>
					<% Else  %>
					<option value="S">Simple Input</option>
					<% End If %>
					<% If rs(fld.name) = "T" then %>
					<option value="T" SELECTED>Multi Line</option>
					<% Else  %>
					<option value="T" >Multi Line</option>
					<% End If %>
					<% If rs(fld.name) = "C" then %>
					<option value="C" SELECTED>Check Box</option>
					<% Else  %>
					<option value="C" >Check Box</option>
					<% End If %>
					</select></td>
<%					end if
					response.write "</tr>"
				end if
			next
			
		%>
		</font>
			</tr>
			</table>
				<br>	
					<input type="submit" name="submit1" value="  Save  " style="font-weight: bold; font-size: x-small; color: White; background-color: #7b68ee;">
					</form>

<p><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>" >
<b>USR_LABEL</b> as it's name suggests this is what the User sees<br>
<b>USR_FIELDTYPE</b> basic dislay control Use either..<br><br></p>
<br>
<p><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>" >
<b>USR_SHORTNAME</b> This is used internally as the forms input name, to retreive the User values.<br>
Try using values like "user_1, user_2" etc, that way it won't clash with any internal names<br><br>
<b>(DO NOT USE SPACES IN THE SHORTNAME)</b></font></p>
<div align="center">
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" >
<a href="../admin_user_fields.asp?uTable=FORUM_USERFIELDS">Back to Table Editor</a>
</font>
			</div>
</body>
</html>