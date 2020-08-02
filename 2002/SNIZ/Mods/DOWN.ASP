<!--#INCLUDE FILE="../config.asp" -->
<!--#INCLUDE FILE="../inc_functions.asp" -->

<html>
<head>
<title>Out of Service Mod - by - AnimeDJ.com</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FFFFFF">

<%
	Dim status, mLev, info1, info2, my_Conn, fStatus

    strCookieURL = Left(Request.ServerVariables("Path_Info"), InstrRev(Request.ServerVariables("Path_Info"), "/Mods/"))
	status = Application("down")
	fStatus = request.form("status")
	DMessage = request.Form("DownMessage")
	if DMessage = "" then
		DMessage = Application("DownMessage")
	end if
	set my_Conn = Server.CreateObject("ADODB.Connection")
	my_Conn.Open strConnString

	mLev = cint(ChkUser2(Request.Cookies(strUniqueID & "User")("Name"), Request.Cookies(strUniqueID & "User")("Pword")))

	my_Conn.close
	set my_Conn = nothing

	if not status then
		status = false
	end if

	if (not isEmpty(fStatus)) and (mLev = 4) then 
		if status then
			Application.lock
			Application("down") = false
			Application("DownMessage") = DMessage
			Application.unlock
			status = false
		else
			Application.lock
			Application("down") = true
			Application("DownMessage") = DMessage
			Application.unlock
			status = true
		end if
	end if

	if status then
		info1 = "down"
		info2 = "Start"
	else
		info1 = "running"
		info2 = "Stop"
	end if

	if mLev = 4 then
%>

<form action="down.asp" method="post">
  <table width="600" border="0" cellspacing="0" cellpadding="10" align="center">
    <tr align="center"> 
      <td> 
        <p><font face="宋体, Arial, Helvetica, sans-serif" size="2"><b>Welcome 
          Administrator. The current status of the boards is <font color="#FF0000"><%= info1 %></font>.</b></font></p>
        <p><font face="宋体, Arial, Helvetica, sans-serif" size="2">would you 
          like to : </font></p>
 
          <input type="submit" value="<%= info2 %> the board" name="Submit">
          <input type="button" value="go back" onClick="location='../admin_home.asp'">

		  <input type="hidden" name="status" value="<%= status %>">

    </td>
  </tr>
  <tr><td align="center">
<p>  <font color="Blue" face="宋体, Arial, Helvetica, sans-serif" size="2">The message below will appear when the board is closed..</font></p>

  <textarea cols="80" rows="12" name="DownMessage" wrap="soft">
<%= Application("DownMessage") %>
</textarea>
</table>
</form>

<%	end if %>

<p>&nbsp;</p>
<p>&nbsp;</p>
<table border="0" cellspacing="0" cellpadding="20" align="center">
  <tr align="center"> 
    <td> 
      <h2><font face="Arial, Helvetica, sans-serif"><%= DMessage %></font></h2>
      </td>
  </tr>
</table>
</body>
</html>
