<table bgcolor="<%=strTableBorderColor%>" cellspacing="1" cellpadding="2" width="95%">
<%			Select Case Request("action")
			Case "updatemodlist"	'## Update Moderator Table for the Forum 

				
				'## Need to verify that Forum ID is present & is a Number
				Dim intRfForumID : intRfForumID = Request.Form("forumid")
				
				If ((intRfForumID <> "") AND IsNumeric(intRfForumID)) Then

					intRfForumID = CInt(intRfForumID)
	
					'## Update Moderator Table
					Dim oCheckBox
					'## 1st Step - clear out current moderators for forum
					'## Forum_SQL - Delete Moderators
					if intRfForumID <> "-1" then
					strSql = "DELETE FROM " & strTablePrefix & "MODERATOR "
					strSql = strSql & " WHERE " & strTablePrefix & "MODERATOR.FORUM_ID = " & intRfForumID & " "
					my_Conn.Execute strSql
					else
					For Each oCheckBox in Request("moderator")
					strSql = "DELETE FROM " & strTablePrefix & "MODERATOR "
					strSql = strSql & " WHERE " & strTablePrefix & "MODERATOR.MEMBER_ID = " & oCheckBox & " "
					my_Conn.Execute strSql
					Next
					end if
					'## 2nd Step - insert the moderators that were checked
					
					For Each oCheckBox in Request("moderator")
						'## Forum_SQL - Insert Moderators
						strSql = "INSERT INTO " & strTablePrefix & "MODERATOR "
						strSql = strSql & "(FORUM_ID"
						strSql = strSql & ", MEMBER_ID"
						strSql = strSql & ") VALUES (" 
						strSql = strSql & intRfForumID
						strSql = strSql & ", " & oCheckBox
						strSql = strSql & ")"
						my_Conn.Execute strSql
					Next


					response.write ("<br>")
					If (strModAdminForumName <> "") Then
						response.write ("<b><a href=""admin_moderators.asp?forum=" & intRfForumID & """>" & strModAdminForumName & "</a></b>�İ��������Ѹ��³ɹ���")
					Else
						response.write ("��ѡ����̳���������Ѹ��£�")
					End If
					response.write ("<form name=""moderatorlist"" method=""post"" action="""">")
					response.write ("<input type=""hidden"" name=""UserID"" value=""a"">")
					response.write ("<input type=""hidden"" name=""forumid"" value="""">")
					response.write "<input type=""hidden"" name=""action"" value="""">" &_
						"<input type=""submit"" name=""Update"" value="" �ص������������� ""></form> "

			
				Else
				
					'## Error with getting the Forum ID
					response.write ("<p>����ʱ��������  Forum ID ���Ϸ���</p>")
				
				End If
			
			Case Else
			
				if Request("UserID") = "" then%>
<tr><td bgcolor="<%=strHeadCellColor%>"><font color="<%=strHeadFontColor%>" size="<%=strDefaultFontSize%>" face="<%=strDefaultFontFace%>">ѡ�����</font></td></tr>
<%					Response.Write ("<tr><td bgcolor=""" & strAltForumCellColor & """><font color=""" & strDefaultFontColor & """ size=""" & strDefaultFontSize & """ face=""" & strDefaultFontFace & """><b>���ࣺ</b> " & strModAdminCategoryName & "</td></tr>" & vbCrLf)
					Response.Write ("<tr><td bgcolor=""" & strAltForumCellColor & """><font color=""" & strDefaultFontColor & """ size=""" & strDefaultFontSize & """ face=""" & strDefaultFontFace & """><b>��̳��</b> " & strModAdminForumName & "</td></tr>" & vbCrLf)%>
<tr><td bgcolor="<%=strHeadCellColor%>"><font color="<%=strHeadFontColor%>" size="<%=strDefaultFontSize%>" face="<%=strDefaultFontFace%>">�������ѡ��</font></td></tr>
<tr>
<td bgcolor="<%=strForumCellColor%>" onmouseover="bgColor='<%=strAltForumCellColor%>'" onmouseout="bgColor='<%=strForumCellColor%>'"><font size="<%=strDefaultFontSize%>">

<%					
					'## Forum_SQL
					strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME "
					strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
					strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_LEVEL > 1 "
					strSql = strSql & " AND   " & strMemberTablePrefix & "MEMBERS.M_STATUS = " & 1
					strSql = strSql & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_NAME ASC;"

					set rs = my_Conn.Execute(strSql)

					response.write ("<form name=""moderatorlist"" method=""post"" action="""">")
					response.write ("<input type=""hidden"" name=""action"" value=""updatemodlist"">")
					response.write ("<input type=""hidden"" name=""forumid"" value=""" & intQsForumID & """>")

					do until rs.EOF

						strRsMember = rs("M_NAME")
						intRsMemberID = rs("MEMBER_ID")

						'## Build list of available moderators					
						response.write ("<input type=""checkbox"" name=""moderator"" value=""" & intRsMemberID & """")
						'## Check to see if already a moderator for this forum
						if chkForumModerator(intQsForumID, strRsMember) then response.write(" ��ѡ��")
						response.write ("> " & strRsMember & "<br>")
						
						rs.MoveNext
					loop
					
					response.write ("<br>" &_
						"<input type=""submit"" name=""Update"" value=""����""> " &_
						"<input type=""reset"" name=""Reset"" value=""����""> " &_
						"</form>")

				else%>
<tr><td bgcolor="<%=strHeadCellColor%>"><font color="<%=strHeadFontColor%>" size="<%=strDefaultFontSize%>" face="<%=strDefaultFontFace%>">ѡ�����</font></td></tr>
<tr>
<td bgcolor="<%=strForumCellColor%>" onmouseover="bgColor='<%=strAltForumCellColor%>'" onmouseout="bgColor='<%=strForumCellColor%>'"><font size="<%=strDefaultFontSize%>">

<%					response.write ("<p>Please select a forum from the left.</p>")
				end if
		End Select
%>
</td>
</tr>
</table>