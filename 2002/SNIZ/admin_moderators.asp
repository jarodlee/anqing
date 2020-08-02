<%
'########## Snitz Forums 2000 Version 3.1 SR4 ####################
'#                                                               #
'#  汉化修改: 资源搜罗站                                         #
'#  电子邮件: cgier@21cn.com                                     #
'#  主页地址: http://www.sdsea.com                               #
'#            http://www.99ss.net                                #
'#            http://www.cdown.net                               #
'#	     http://www.wzdown.com                               #
'#	     http://www.13888.net                                #
'#  论坛地址:http://ubb.yesky.net                                #
'#  最后修改日期: 2001/03/12    中文版本：Version 3.1 SR4        #
'#################################################################
'# 原始来源                                                      #
'# Snitz Forums 2000 Version 3.1 SR4                             #
'# Copyright 2000 http://forum.snitz.com - All Rights Reserved   #
'#################################################################
'#【版权声明】                                                   #
'#                                                               #
'# 本软体为共享软体(shareware)提供个人网站免费使用，请勿非法修改,#
'# 转载，散播，或用于其他图利行为，并请勿删除版权声明。          #
'# 如果您的网站正式起用了这个脚本，请您通知我们，以便我们能够知晓#
'# 如果可能，请在您的网站做上我们的链接，希望能给予合作。谢谢！  #
'#################################################################
'# 请您尊重我们的劳动和版权，不要删除以上的版权声明部分，谢谢合作#
'# 如有任何问题请到我们的论坛告诉我们                            #
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
    <img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="default.asp">返回论坛首页</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="admin_home.asp">论坛管理中心</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" height=15 width=15 border="0">&nbsp;斑竹设置<br>
    </font></td>
  </tr>
</table>
<%
	if request("Forum") = "" then
		txtMessage = "请任意选择一个论坛修改该论坛版主"
	else
		if request("userid") = "" then
			'txtMessage = "选择一名会员赋予或取消版主权力。粗黑体表示为现任版主"
			
			'## My message  :)
			txtMessage = "请任意选择一个论坛编辑该论坛版主。前面打勾为现任版主"
		else
			if Request("action") = "" then
				txtMessage = "修改此会员权力"
			else
				txtMessage = "修改成功！"
			end if
		end if
	end if
%>
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td bgcolor="<% =strTableBorderColor %>">
<table border="0" width="100%" cellspacing="1" cellpadding="4">
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" size"<% =strDefaultFontSize %>">斑竹设置<%if txtMessage <> "" Then%><br><%=txtMessage%><%End If%></font></td>
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
        <tr><td bgcolor="<%=strForumCellColor%>"><font color="<%=strHeadFontColor%>" size="<%=strDefaultFontSize%>" face="<%=strDefaultFontFace%>"><a href="admin_moderators.asp?forum=-1">所有论坛板块</a></font></td></tr>
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
        response.write("没有任何论坛<br>")
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
<tr><td width="100%" bgcolor="<%=strHeadCellColor%>"><font color="<%=strHeadFontColor%>" size="<%=strDefaultFontSize%>" face="<%=strDefaultFontFace%>">选择版主</font></td></tr>
<tr>
<td bgcolor="<%=strForumCellColor%>" onmouseover="bgColor='<%=strAltForumCellColor%>'" onmouseout="bgColor='<%=strForumCellColor%>'"><font size="<%=strDefaultFontSize%>">

<%		response.write ("<p>请从论坛列表中选择一个板块。</p></td></tr></table>")
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
