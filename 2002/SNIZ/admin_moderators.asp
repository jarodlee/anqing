<%
'########## Snitz Forums 2000 Version 3.1 SR4 ####################
'#                                                               #
'#  �����޸�: ��Դ����վ                                         #
'#  �����ʼ�: cgier@21cn.com                                     #
'#  ��ҳ��ַ: http://www.sdsea.com                               #
'#            http://www.99ss.net                                #
'#            http://www.cdown.net                               #
'#	     http://www.wzdown.com                               #
'#	     http://www.13888.net                                #
'#  ��̳��ַ:http://ubb.yesky.net                                #
'#  ����޸�����: 2001/03/12    ���İ汾��Version 3.1 SR4        #
'#################################################################
'# ԭʼ��Դ                                                      #
'# Snitz Forums 2000 Version 3.1 SR4                             #
'# Copyright 2000 http://forum.snitz.com - All Rights Reserved   #
'#################################################################
'#����Ȩ������                                                   #
'#                                                               #
'# ������Ϊ��������(shareware)�ṩ������վ���ʹ�ã�����Ƿ��޸�,#
'# ת�أ�ɢ��������������ͼ����Ϊ��������ɾ����Ȩ������          #
'# ���������վ��ʽ����������ű�������֪ͨ���ǣ��Ա������ܹ�֪��#
'# ������ܣ�����������վ�������ǵ����ӣ�ϣ���ܸ��������лл��  #
'#################################################################
'# �����������ǵ��Ͷ��Ͱ�Ȩ����Ҫɾ�����ϵİ�Ȩ�������֣�лл����#
'# �����κ������뵽���ǵ���̳��������                            #
'#################################################################
%>
<!--#INCLUDE FILE="config.asp"-->
<% if Session(strCookieURL & "Approval") = "15916941253" then %>
<!--#INCLUDE FILE="inc_functions.asp"-->
<!--#INCLUDE FILE="inc_top.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<table border="0" width="100%">
  <tr>
    <td width="33%" align="left" nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    <img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="default.asp">������̳��ҳ</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="admin_home.asp">��̳��������</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" height=15 width=15 border="0">&nbsp;��������<br>
    </font></td>
  </tr>
</table>
<%
	if request("Forum") = "" then
		txtMessage = "������ѡ��һ����̳�޸ĸ���̳����"
	else
		if request("userid") = "" then
			'txtMessage = "ѡ��һ����Ա�����ȡ������Ȩ�����ֺ����ʾΪ���ΰ���"
			
			'## My message  :)
			txtMessage = "������ѡ��һ����̳�༭����̳������ǰ���Ϊ���ΰ���"
		else
			if Request("action") = "" then
				txtMessage = "�޸Ĵ˻�ԱȨ��"
			else
				txtMessage = "�޸ĳɹ���"
			end if
		end if
	end if
%>
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td bgcolor="<% =strTableBorderColor %>">
<table border="0" width="100%" cellspacing="1" cellpadding="4">
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" size"<% =strDefaultFontSize %>">��������<%if txtMessage <> "" Then%><br><%=txtMessage%><%End If%></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
<%	'if Request("Forum") = "" then %>

<table border="0" cellspacing="1" cellpadding="3" align="left" width=100%>
  <tr valign="top">
	<td width="40%">
<table bgcolor="<%=strTableBorderColor%>" cellspacing="1" cellpadding="3">
<%

		Dim strModAdminCategoryName : strModAdminCategoryName = "NA"
		Dim strModAdminForumName : strModAdminForumName = "NA"
		Dim intQsForumID : intQsForumID = Request("Forum")
		Dim intRsForumID : intRsForumID = ""
		Dim strRsMember, intRsMemberID
		
		If IsNumeric(intQsForumID) Then intQsForumID = CInt(intQsForumID)

		'## Forum_SQL
        '''''''Added by me
        strsql = "Select "&strTablePrefix&"CATEGORY.CAT_ID, "&strTablePrefix&"CATEGORY.CAT_NAME from "&strTablePrefix&"CATEGORY order by "&strTablePrefix&"CATEGORY.CAT_NAME asc"
		set drs = my_conn.execute(strsql)%>
        <tr><td bgcolor="<%=strForumCellColor%>"><font color="<%=strHeadFontColor%>" size="<%=strDefaultFontSize%>" face="<%=strDefaultFontFace%>"><a href="admin_moderators.asp?forum=-1">������̳���</a></font></td></tr>
        <%Do until drs.eof
        strAddedCat_Name = drs("CAT_NAME")
        strAddedCat_Id   = drs("CAT_ID")
        %>
		<tr><td  bgcolor="<%=strHeadCellColor%>"><font color="<%=strHeadFontColor%>" size="<%=strDefaultFontSize%>" face="<%=strDefaultFontFace%>"><%=strAddedCat_Name%></font></td></tr>

        <%
		strSql = "SELECT " & strTablePrefix & "FORUM.CAT_ID, " & strTablePrefix & "FORUM.FORUM_ID, " & strTablePrefix & "FORUM.F_SUBJECT "
		strSql = strSql & " FROM " & strTablePrefix & "FORUM  WHERE "&strTablePrefix&"FORUM.CAT_ID="&strAddedCat_Id
		strSql = strSql & " ORDER BY " & strTablePrefix & "FORUM.CAT_ID ASC, " & strTablePrefix & "FORUM.F_SUBJECT ASC;"

		set rs = my_Conn.Execute(strSql)
		if drs.eof and drs.bof and rs.EOF and rs.BOF then
        response.write("û���κ���̳<br>")
        else
			do until rs.EOF
				intRsForumID = rs.Fields("FORUM_ID").Value
				If IsNumeric(intRsForumID) Then intRsForumID = CInt(intRsForumID)

				'## Check if Querystring and Recordset ForumID's are the same
				'## If true, then assign category and forum titles to variables
				If (intQsForumID = intRsForumID) Then
					strModAdminCategoryName = strAddedCat_Name
					strModAdminForumName = rs.Fields("F_SUBJECT").Value
				End If
				If (intQsForumID = -1) Then
					strModAdminCategoryName = "All Categories"
					strModAdminForumName = "All Forums"
				End If
				%>
				<tr><td bgcolor="<%=strForumCellColor%>" onmouseover="bgColor='<%=strAltForumCellColor%>'" onmouseout="bgColor='<%=strForumCellColor%>'"><font size="<%=strDefaultFontSize%>"><a href="admin_moderators.asp?forum=<%=rs("FORUM_ID")%>"><%=rs("F_SUBJECT")%></a></td></tr>
				<%
				rs.MoveNext
			loop
        end if
        ''added
		drs.movenext
		loop
		set drs = nothing
		''Done changing
		%>
		</dl>
		</td></tr></table></td>
		<td>&nbsp;</td>
		<td><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
		<%

	'else
	if (Request("Forum") <> "") AND IsNumeric(Request("Forum")) then%>
<!--#include file="update_moderator.asp"-->
<%	else%>
<table bgcolor="<%=strTableBorderColor%>" cellspacing="1" cellpadding="2" width="100%">
<tr><td width="100%" bgcolor="<%=strHeadCellColor%>"><font color="<%=strHeadFontColor%>" size="<%=strDefaultFontSize%>" face="<%=strDefaultFontFace%>">ѡ�����</font></td></tr>
<tr>
<td bgcolor="<%=strForumCellColor%>" onmouseover="bgColor='<%=strAltForumCellColor%>'" onmouseout="bgColor='<%=strForumCellColor%>'"><font size="<%=strDefaultFontSize%>">

<%		response.write ("<p>�����̳�б���ѡ��һ����顣</p></td></tr></table>")
	end if
%>
    </font>
    </td></tr></table>
    </font></td>
  </tr>
</table>
    </td>
  </tr>
</table>
<!--#INCLUDE FILE="inc_footer.asp"-->
<% else %>
<%	Response.Redirect "admin_login.asp" %>
<% end iF %>
