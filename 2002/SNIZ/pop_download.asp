<!--#INCLUDE FILE="config.asp" --> 
<!--#INCLUDE FILE="inc_functions.asp" -->
<!--#INCLUDE FILE="inc_top_short.asp" -->

<% select case Request.QueryString("mode")  %>

<% case "Edit" %>

	<center>
	<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">�ļ�����</font></p>
	<p align=center>
	<form action="pop_download.asp?mode=goEdit&dir=<% =Request.QueryString("dir")%>&file=<% =Request.QueryString("file")%>" method="post">
	<input name="Refer" type="hidden" value="<% =Request.ServerVariables("HTTP_REFERER") %>">
	<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">ֻ�б���̳ע���û����������ļ���<br>

<%		if strAuthType = "nt" then %>
��� NT ʹ�����ʺ�����ʾ����ѡ���ͳ���������<br><br>
<%		else %>
<%			if strAuthType = "db" then %>
����д���<br><br>
<%			end if %>	
<%		end if %>	

	�������δע�ᣬ<a href="policy.asp">������ע��</a>��</center>

	<table border="0" cellspacing="0" cellpadding="0" align=center>
		<tr>
			<td bgcolor="<% =strPopUpBorderColor %>">
	<table border="0" width="100%" cellspacing="1" cellpadding="1">
<%		if strAuthType = "nt" then %>
		<TR>
			<TD bgColor="<% =strPopUpTableColor %>" align=right nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">����ʺ�</font></b></td>
			<TD bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><% =Session(strCookieURL & "userid") %></font></b></td>
		</TR>
<%		else %>	
<%			if strAuthType = "db" then %>
	 	<TR>
	 	    <TD  bgColor=<% =strPopUpTableColor %> align=right nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�û�����</font></b></td>
	 	    <TD  bgColor=<% =strPopUpTableColor %>><input name="Name" size="25" value="<% =Request.Cookies(strUniqueID & "User")("Name")%>"></td>
	 	</TR>
	 	<TR>
	 	    <TD  bgColor=<% =strPopUpTableColor %> align=right nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">���룺</font></b></td>
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
	<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�㲢�ǻ�Ա�����ع���δ���š�<br></font>
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

	<p><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">ѡ�����������������ļ�</font></p><br>
	<% strFileLink = strRoot & Request("dir") & "/" & fileName %>
	<a href="<%= strFileLink %>" target="_self" ><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b><%= fileName %></b></font></a><br>
</center>
<br>
<br>
<br>
<br>
    <p><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript:onClick= window.close()">�رմ���</A></font></p>

<%
	response.end
end if
rs.close
set rs=nothing

%>

<%
end select


%>

