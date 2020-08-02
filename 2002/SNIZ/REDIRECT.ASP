<% 
Dim UseCount,URL_REDIR,URLid

function remoteIP()	
	
	remoteIP = Request.ServerVariables("HTTP_X_FORWARDED_FOR")	
	if remoteIP = "" then		
		remoteIP = Request.ServerVariables("REMOTE_ADDR")	
	end if
end function


strConnString = "driver={SQL Server};server=127.0.0.1;uid=sa;pwd=;database=VIRTUE_LOAD" '## MS SQL Server 7
set my_Conn= Server.CreateObject("ADODB.Connection")
my_Conn.Open strConnString

strSQL = "SELECT ID, UseageCount, URLPath "
strSQL = strSQL & "FROM PAGE_REDIRECT WHERE UseageCount = (SELECT MIN(UseageCount) FROM PAGE_REDIRECT)"


set rsData =  my_Conn.Execute (strSql)
if not rsData.EOF then
	UseCount = rsData("UseageCount")
	URL_REDIR = rsData("URLPath")
	URLid = rsData("ID")
	
	strSQL = "UPDATE PAGE_REDIRECT SET UseageCount = UseageCount + 1 WHERE ID = " & URLid
	my_conn.Execute(strSQL)
	strSQL = "INSERT  INTO REDIRECT_LOG (URLid,UserIP) VALUES ("
	strSQL = strSQL & URLid & ",'" & remoteIP & "')"
	my_Conn.Execute (strSql)
	
	
	set rsData = nothing
	my_Conn.Close
	set my_conn = nothing
	
	response.write "Redirecting to " & URL_REDIR
 %>
 <meta http-equiv="Refresh" content="1; URL=<% =URL_REDIR%>"> 
<% Else  
	if Request.Form("inURLPath") = "" then
		Response.Write "Either No URL's are defined, or All ports are Blocked "%>
		<form action="redirect.asp" method="post" name="URLUpdate" id="URLUpdate">
		Name &nbsp;&nbsp;&nbsp;<input type="text" name="inName" value="" maxlength="50"><br>
		New URL &nbsp;&nbsp;&nbsp;<input type="text" name="inURLPath" value="http://" maxlength="500"><br>
		IP Address &nbsp;<input type="text" name="inIP" value="000.000.000.000" maxlength="20"><br>		
		
		Password&nbsp;&nbsp;<input type="text" name="password" value="" maxlength="20"><br>
		<input type="hidden" name="pwd" value="virtue2k">
		<input type="submit" name="submit" value="Add URL">
		</form>
	<% Else 
		 if Request.Form("password") = Request.Form("pwd") then
		 strSQL = "INSERT  INTO PAGE_REDIRECT (URLPath,IP,Name) VALUES ('" & Request.Form("inURLPath") & "," & Request.Form("inIP") & "," & Request.Form("inName") &"')" 
		 my_conn.execute strSQL
		 Response.Write "Added new URL<br>" %>
		  <meta http-equiv="Refresh" content="1; URL=redirect.asp"> 
<%		end if 
		end if%>
<% End If %>