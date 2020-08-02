<%	Dim Forum_Count
	Dim NewMember_Name, NewMember_Id, Member_Count
	Dim LastPostDate, LastPostLink

	set rs = Server.CreateObject("ADODB.Recordset")
	
	Forum_Count = intForumCount

	'## Forum_SQL - Get newest membername and id from DB
	strSql = "SELECT M_NAME, MEMBER_ID FROM " & strMemberTablePrefix & "MEMBERS WHERE M_STATUS=1 AND MEMBER_ID > 1"
	strSql = strSQL & " ORDER BY M_DATE desc;"
	set rs = my_Conn.Execute(strSql)
	if not rs.EOF then
		NewMember_Name = ChkString(rs("M_NAME"), "display") 
		NewMember_Id = rs("MEMBER_ID")
	else 
		NewMember_Name = ""
	end if
    
	'## Forum_SQL - Get Active membercount from DB 
	strSql = "SELECT COUNT(MEMBER_ID) AS U_COUNT FROM " & strMemberTablePrefix & "MEMBERS WHERE M_POSTS > 0 AND M_STATUS=1"
	
	
	set rs = my_Conn.Execute(strSql)
	
	if not rs.EOF then
		Member_Count = rs("U_COUNT")
	else
		Member_Count = 0
	end if
	
	LastPostDate = ""
 	LastPostLink = ""
	LastPostAuthorLink = ""
	
	if not (intLastPostForum_ID = "") then	
		'## Forum_SQL - Get lastPostDate and link to that post from DB
		strSql = "SELECT " & strTablePrefix & "FORUM.CAT_ID, " & strTablePrefix & "FORUM.FORUM_ID, " 
		strSql = strSql & strTablePrefix & "FORUM.F_SUBJECT, " & strTablePrefix & "TOPICS.TOPIC_ID, " & strTablePrefix & "TOPICS.T_SUBJECT, "
		strSql = strSql & strTablePrefix & "TOPICS.T_LAST_POST, " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.MEMBER_ID "
		strSql = strSql & "FROM " & strTablePrefix & "FORUM, " & strTablePrefix & "TOPICS, "
		strSql = strSql & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " WHERE " & strTablePrefix & "FORUM.FORUM_ID = " & strTablePrefix & "TOPICS.FORUM_ID "
		strSql = strSql & " AND " & strTablePrefix & "FORUM.CAT_ID = " & strTablePrefix & "TOPICS.CAT_ID "
		strSql = strSql & " AND " & strTablePrefix & "TOPICS.T_LAST_POST_AUTHOR = " & strMemberTablePrefix & "MEMBERS.MEMBER_ID "
		strSql = strSql & " AND " & strTablePrefix & "FORUM.FORUM_ID = " & intLastPostForum_ID & " "
		strSql = strSql & "ORDER BY " & strTablePrefix & "TOPICS.T_LAST_POST DESC;"

	 	set rs = my_Conn.Execute(strSql)
 	
 		if not rs.EOF then
			LastPostDate = ChkDate(rs("T_LAST_POST")) & ChkTime(rs("T_LAST_POST"))
			LastPostLink = "topic.asp?TOPIC_ID=" & rs("TOPIC_ID") & "&FORUM_ID=" & rs("FORUM_ID") & "&CAT_ID=" & rs("CAT_ID")
			LastPostLink = LastPostLink  & "&Topic_Title=" & ChkString(rs("T_SUBJECT"),"urlpath")
			LastPostLink = LastPostLink  & "&Forum_Title=" & ChkString(rs("F_SUBJECT"),"urlpath") 
			LastPostAuthorLink = " ���ߣ�"
			strMember_ID = rs("MEMBER_ID")
			strM_NAME = ChkString(rs("M_NAME"),"display") 
			if strUseExtendedProfile then
				LastPostAuthorLink = LastPostAuthorLink & "<a href=""pop_profile.asp?mode=display&id="& strMember_ID & """>"
			else
				LastPostAuthorLink = LastPostAuthorLink & "<a href=""JavaScript:openWindow2('pop_profile.asp?mode=display&id=" & strMember_ID & "')"">"
			end if
            LastPostAuthorLink = LastPostAuthorLink  & strM_NAME & "</a>"
		end if
	end if

	ActiveTopicCount = -1
	if not IsNull(Session(strCookieURL & "last_here_date")) then 
		if not blnHiddenForums then
			'## Forum_SQL - Get ActiveTopicCount from DB
			strSql = "SELECT COUNT(" & strTablePrefix & "TOPICS.T_LAST_POST) AS NUM_ACTIVE "
			strSql = strSql & "FROM " & strTablePrefix & "TOPICS "
			strSql = strSql & "WHERE (((" & strTablePrefix & "TOPICS.T_LAST_POST)>'"& Session(strCookieURL & "last_here_date") & "'))"

			set rs = my_Conn.Execute(strSql)
			if not rs.EOF then
				ActiveTopicCount = rs("NUM_ACTIVE")
			else
				ActiveTopicCount = 0
			end if
		end if
	end if

	rs.close
	set rs = nothing
	ShowLastHere = (cint(ChkUser2(strDBNTUserName, Request.Cookies(strUniqueID & "User")("Pword"))) > 0)%>
          <td bgcolor="<% =strCategoryCellColor %>" colspan="<% if (strShowModerators = "1") or (mlev = 4 or mlev = 3) then Response.Write("6") else Response.Write("5")%>">
<%		' This code will specify whether or not to show the forums under a category

		HideForumStat = "HideCat" & 0
		If Request.Cookies(HideForumStat) = "Y" then %>
<%			if (InStr(1, ScriptName, "default.asp", 1)) then %>
			        <a href="<% =ScriptName & "?" %><% =HideForumStat & "=N" %>"><img src="<%=strImageURL %>icon_plus.gif" width="10" height="10" border="0"></a>
<%			else %>
				<a href="<% =ScriptName & "default.asp?" %><% =HideForumStat & "=N" %>"><img src="<%=strImageURL %>icon_plus.gif" width="10" height="10" border="0"></a>
<%			end if %>
<%		Else %>
<%			if (InStr(1, ScriptName, "default.asp", 1)) then %>
				<a href="<% =ScriptName & "?" %><% =HideForumStat & "=Y" %>"><img src="<%=strImageURL %>icon_minus.gif" width="10" height="10" border="0"></a>
<%			else %>
				<a href="<% =ScriptName & "default.asp?" %><% =HideForumStat & "=Y" %>"><img src="<%=strImageURL %>icon_minus.gif" width="10" height="10" border="0"></a>
<%			end if %>
<%		end if %>
	  <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" size="+1"><b>״̬��</b></font></td>
        </tr>



<%     if Request.Cookies(HideForumStat) <> "Y" then %>
      <tr>
        <td rowspan="<% if (ShowLastHere and NewMember_Name <> "") then Response.Write("4") else Response.Write("3") end if %>" bgcolor="<%= strForumCellColor %>">&nbsp;</td>
<%
	if ShowLastHere then 
%>
        <td bgcolor="<%= strForumCellColor %>" colspan="<% if ((strShowModerators = "1") or (mlev = "4" or mlev = "3")) then Response.Write("6") else Response.Write("4") end if %>">
        <font align=left face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">ǰ�ι���ʱ�䣺<% =ChkDate(Session(strCookieURL & "last_here_date")) %> <% =ChkTime(Session(strCookieURL & "last_here_date")) %></font>
        </td>
	  </tr>
	  <tr>
<%
	end if 
	if intPostCount > 0 then 
%>
        <td bgcolor="<%= strForumCellColor %>" colspan="<% if ((strShowModerators = "1") or (mlev = "4" or mlev = "3")) then Response.Write("6") else Response.Write("4") end if%>">
        <font face="<%= strDefaultFontFace %>" size="<% =strFooterFontSize %>">
		<% 
		strUser = Users & " ����Ա"
		if Member_Count = 1 then 
				if users = 1 then
					Response.Write(strUser & "�������� ") 
				else
					Response.Write(strUser & "�е� 1 λ�������� ") 
				end if
		else 
				Response.Write(strUser & "�е� " & Member_Count & " λ�������� ") 
		end if 
		Response.Write(intPostCount & " ƪ������ ")  
		Response.Write(intForumCount & " ��������") 

		if (LastPostDate = "" or LastPostLink = "" or intPostCount = 0) then 
			Response.Write("��") 
		else
			Response.Write("�������·������������ <a href=""" & lastPostLink & """>"& lastPostDate & "</a>")
			if  LastPostAuthorLink <> "" then
				Response.Write(LastPostAuthorLink & "��")
			else
				Response.Write("��")
			end if
		end if
%>
		  </font>
          </td>
        </tr>

        <tr>
<%
	end if
%>      
        <td bgcolor="<%= strForumCellColor %>" colspan="<% if ((strShowModerators = "1") or (mlev = "4" or mlev = "3")) then Response.Write("6") else Response.Write("4") end if%>">
        <font face="<%= strDefaultFontFace %>" size="<% =strFooterFontSize %>">����̳���� <%= intTopicCount %> ƪ����<% if ActiveTopicCount > 0 then %>�������� <%= ActiveTopicCount %> ��<a href="active.asp">������������</a>��<% elseif blnHiddenForums and (strLastPostDate > Session(strCookieURL & "last_here_date")) and ShowLastHere then %>��������<a href="active.asp">������������</a>��<% elseif not(ShowLastHere) then Response.Write "��" else %>�����ǲ�û�������¡�<% end if %></font>
        </td>
      </tr>
<%
	if NewMember_Name <> "" then 
%>
      <tr>          
        <td bgcolor="<%= strForumCellColor %>" colspan="<% if ((strShowModerators = "1") or (mlev = "4" or mlev = "3")) then Response.Write("6") else Response.Write("4") end if%>">
        <font face="<%= strDefaultFontFace %>" size="<% =strFooterFontSize %>">���һ�ӭ���ǵ��»�Ա��
<%
		if strUseExtendedProfile then
			Response.Write "        <a href=""pop_profile.asp?mode=display&id="& NewMember_Id & """>"
		else
			Response.Write "        <a href=""JavaScript:openWindow2('pop_profile.asp?mode=display&id=" & NewMember_Id & "')"">"
		end if
		Response.Write NewMember_Name & "</a>.</font>" & vbcrlf
%>
          </td>
        </tr>
<%
	end if 
End If %>
