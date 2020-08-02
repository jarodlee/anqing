<!--#INCLUDE FILE="config.asp" --> 
<!--#INCLUDE FILE="inc_functions.asp" -->
<!--#INCLUDE FILE="inc_top_short.asp" -->

<% select case Request.QueryString("mode")  %>

<% case "Edit" %>

	<center>
	<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">文件下载</font></p>
	<p align=center>
	<form action="pop_download.asp?mode=goEdit&dir=<% =Request.QueryString("dir")%>&file=<% =Request.QueryString("file")%>" method="post">
	<input name="Refer" type="hidden" value="<% =Request.ServerVariables("HTTP_REFERER") %>">
	<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">只有本论坛注册用户可以下载文件。<br>

<%		if strAuthType = "nt" then %>
你的 NT 使用者帐号已显示。点选『送出』继续。<br><br>
<%		else %>
<%			if strAuthType = "db" then %>
请填写表格。<br><br>
<%			end if %>	
<%		end if %>	

	如果你尚未注册，<a href="policy.asp">按这里注册</a>。</center>

	<table border="0" cellspacing="0" cellpadding="0" align=center>
		<tr>
			<td bgcolor="<% =strPopUpBorderColor %>">
	<table border="0" width="100%" cellspacing="1" cellpadding="1">
<%		if strAuthType = "nt" then %>
		<TR>
			<TD bgColor="<% =strPopUpTableColor %>" align=right nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">你的帐号</font></b></td>
			<TD bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><% =Session(strCookieURL & "userid") %></font></b></td>
		</TR>
<%		else %>	
<%			if strAuthType = "db" then %>
	 	<TR>
	 	    <TD  bgColor=<% =strPopUpTableColor %> align=right nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">用户名：</font></b></td>
	 	    <TD  bgColor=<% =strPopUpTableColor %>><input name="Name" size="25" value="<% =Request.Cookies(strUniqueID & "User")("Name")%>"></td>
	 	</TR>
	 	<TR>
	 	    <TD  bgColor=<% =strPopUpTableColor %> align=right nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">密码：</font></b></td>
	 	    <TD  bgColor=<% =strPopUpTableColor %>><input name="Password" type=Password size="25" value="<% =Request.Cookies(strUniqueID & "User")("Pword")%>"></td>
	 	</TR>
<%			end if %>
<%		end if %>
	 	<TR>
	 		<TD  bgColor=<% =strPopUpTableColor %> align=center colspan=2><input type=submit value=Submit></td>
	 	</TR>    
	</table>
			</td>
		</tr>
	</table>
	
	</form></p>

<% 	case "goEdit" 
		
		strSql = "SELECT M_ALLOWDOWNLOADS"
		strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " WHERE "&Strdbntsqlname&" = '" & Request.Form("Name") & "' "
		if strAuthType = "db" then
				strSql = strSql & " AND   M_PASSWORD = '" & ChkString(Request.Form("Password"),"password") & "'"
		end if
		fileName = Request("file")
		pos = Instr(filename,"target")
		if pos <> 0 then
		filename = Left(filename,pos-2)
		end if

		on error resume next
		set rs = my_Conn.Execute (strSql)

		if  (Request.Form("Name") = "") or (rs("M_ALLOWDOWNLOADS") <> "1") then

%>
	<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">你并非会员或下载功能未开放。<br></font>
	<meta http-equiv="Refresh" content="2; URL=default.asp">
<center>
<!--#INCLUDE FILE="inc_footer_short.asp" -->
</center>	
<% 
			response.end
		else					
     	strRoot = replace(strCookieURL,"/mods/","/",1,-1,1) & "uploaded/"
%>
<center>

	<p><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">选择下面连接来下载文件</font></p><br>
	<% strFileLink = strRoot & Request("dir") & "/" & fileName %>
	<a href="<%= strFileLink %>" target="_self" ><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b><%= fileName %></b></font></a><br>
</center>
<br>
<br>
<br>
<br>
    <p><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript:onClick= window.close()">关闭窗口</A></font></p>

<%
	response.end
end if
rs.close
set rs=nothing

%>

<%
end select


%>

