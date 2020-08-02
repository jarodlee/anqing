<% server.scripttimeout = 6000 %>
<!--#INCLUDE FILE="config.asp" -->
<% if Session(strCookieURL & "Approval") = "15916941253" then %>
<!--#INCLUDE FILE="inc_functions.asp" -->
<!--#INCLUDE file="inc_top.asp" -->
<script language="JavaScript">
<!-- hide from JavaScript-challenged browsers
function selectAll(formObj, isInverse) 
{ 
with (formObj) 
{ 
for (var i=0;i < length;i++) 
{ 
fldObj = elements[i]; 
if(isInverse) 
{ 
if (fldObj.name != 'inverse') 
{ 
if (fldObj.name == 'selectall') 
fldObj.checked = false; 
else 
fldObj.checked = (fldObj.checked) ? false : true; 
} 
else fldObj.checked = true; 
} 
else 
{ 
fldObj.checked = true; 
if (fldObj.name == 'inverse') fldObj.checked = false; 
} 
} 
} 
}
// done hiding --> 
</script>
<%
if request.form("MSG") = "kerrywrotethis" then
response.redirect "admin_emailmanager.asp"
end if


PageName = request.form("pagename")
if PageName = "" then
My_Sort = request.form("sortme")
My_Order = request.form("order")
My_Order2 = request.form("order2")

if My_Sort = "" then
Sort_Name = "所有会员"
elseif My_Sort = "1" then
Sort_Name = "只有一般会员" 
elseif My_Sort = "2" then
Sort_Name = "只有版主"
elseif My_Sort = "3" then
Sort_Name = "只有管理员"
elseif My_Sort = "4" then
Sort_Name = "半年未使用者"
My_Last = DateToStr(DateAdd("m", -6, now()))
elseif My_Sort = "5" then
Sort_Name = "一年未使用者"
My_Last = DateToStr(DateAdd("yyyy", -1, now()))
elseif My_Sort = "6" then
Sort_Name = "无发表文章者"
end if

if My_Order = "" then
Order_Name = "按会员名"
elseif  My_Order ="2" then
Order_Name = "按等级"
elseif  My_Order ="3" then
Order_Name = "按文章数" 
elseif  My_Order ="4" then
Order_Name = "按注册日期"
end if

if My_Order2 = "" then
Order_By = "Desc"
elseif My_Order2 = "1" then
Order_By = "Asc"
end if
%>
<table border="0" width="100%">
  <tr>
 <td width="33%" align="left" nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
 <img src="images/icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="default.asp">返回论坛首页</a>&nbsp;<img src="images/icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="admin_home.asp">论坛管理中心</a>&nbsp;<img src="images/icon_folder_open_topic.gif" height=15 width=15 border="0">&nbsp;邮件列表设置<br>
 </font></td>
  </tr>
</table>
<center>
<table border="0" cellpadding="10">
  <tr><form method="post" action="admin_emaillist.asp">
 <td nowrap>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">选择发送群组：</font><BR>
 <font face="宋体, Arial, Helvetica" size="2"> 
 <select name="sortme" size="1">
  <option value="<%=My_Sort%>" SELECTED>&nbsp;<% =Sort_Name%></option>
  <option value="">&nbsp;所有会员</option>
  <option value="3">&nbsp;只有管理员</option>
  <option value="2">&nbsp;只有版主</option>
  <option value="1">&nbsp;只有一般会员</option>
  <option value="4">&nbsp;半年未使用者</option>
  <option value="5">&nbsp;一年未使用者</option>
  <option value="6">&nbsp;无发表文章者</option>
  </select>
  </font>
 </td><td>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">排序方式：</font><BR>
 <font face="宋体, Arial, Helvetica" size="2"> 
 <select name="order" size="1">
  <option value="<%=My_Order%>" SELECTED>&nbsp;<%=Order_Name%></option>
  <option value="">&nbsp;按会员名</option>
  <option value="2">&nbsp;按等级</option>
  <option value="3">&nbsp;按文章数</option>
  <option value="4">&nbsp;按注册日期</option>
</select>
 </font>
 </td>
<td>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">排序方法：</font><BR>
 <font face="宋体体, Arial, Helvetica" size="2"> 
 <select name="order2" size="1">
  <option value="<%=My_Order2%>" SELECTED>&nbsp;<%=Order_By%></option>
  <option value="">&nbsp;由大到小</option>
  <option value="1">&nbsp;由小到大</option>
</select>
 <input type="submit" value="显示">
 </font>
 </td></form>
</tr><form action="admin_emaillist.asp" method="post"></table>



<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">选择信件内容：

<%         '------------------------------------
strSql = "SELECT * FROM " &strMemberTablePrefix & "SPAM WHERE ARCHIVE = '0'"
set rsSP = Server.CreateObject("ADODB.Recordset")
rsSP.open  strSql, my_Conn, 3
%>
<% if rsSP.EOF or rsSP.BOF then %>
建立新讯息
<% else %>
<select name="MSG" size="1">
  <option value="" SELECTED>&nbsp;- 建立新讯息- </option>
<%
do until rsSP.EOF 
%>

  <option value="<% =rsSP("ID") %>">&nbsp;<% =Server.HTMLEncode(rsSP("Subject")) %></option>
  
<%
rsSP.MoveNext
%>
<%loop %>
</select>
<% end if 
set rsSP = nothing
%>
 &nbsp;&nbsp; 或 <a href="admin_emailmanager.asp">进入讯息管理中心</a></font>
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center">
  <tr>
  <input type="hidden" name="pagename" value="compose">
  <td align=center>
  <input type="button" value="选择本页全部会员" onClick="selectAll(this.form,1)">&nbsp;
<input type="submit" name="action" value="将讯息寄给所选择的会员">
&nbsp;<input type="submit" name="action" value="将讯息寄给所有会员">&nbsp;<input type="reset" name="Reset" value="全部清除">
<BR><BR></td></tr>
 <td bgcolor="<% =strTableBorderColor %>">
 <table border="0" width="100%" cellspacing="1" cellpadding="4">
<tr align=center>
<td bgColor="<% =strHeadCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">选择</font></td>
  <td bgColor="<% =strHeadCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">会员名</font></td>
  <td bgColor="<% =strHeadCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">邮件地址</font></td>
  <td bgColor="<% =strHeadCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">发表</font></td>
  <td bgColor="<% =strHeadCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">等级</font></td>  
  <td bgColor="<% =strHeadCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">注册日期</font></td>
  <td bgColor="<% =strHeadCellColor %>">&nbsp;</td>  
</tr>
<%
strSql = "SELECT * FROM " &strMemberTablePrefix & "MEMBERS "
strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_STATUS = " & 1
strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_EMAIL <> " & "''"

if My_Sort > "" then
if My_Sort = "1" then
  strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_LEVEL = " & 1
elseif My_Sort = "2" then
 strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_LEVEL = " & 2
elseif My_Sort = "3" then 
 strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_LEVEL = " & 3
elseif My_Sort = "6" then
  strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_POSTS = " & 0  
elseif My_Sort = "4" or "5" then 
 strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_LASTPOSTDATE < '" & My_Last & "'"
 end if
 end if
 
if My_Order = "2" then
 strSql = strSql & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_LEVEL " & Order_By
elseif My_Order = "3" then 
 strSql = strSql & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_POSTS " & Order_By
elseif My_Order = "4" then 
 strSql = strSql & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_DATE " & Order_By
else
 strSql = strSql & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_NAME " & Order_By
 end if
 
set rs = Server.CreateObject("ADODB.Recordset")
rs.cachesize=20
rs.open  strSql, my_Conn, 3
%>
<% if rs.EOF or rs.BOF then %>
<tr>
  <td bgcolor="<% =strForumCellColor %>" colspan="7"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">目前没有任何会员</font></td>
</tr>
<% else %>
<%do until rs.EOF 
MyRank = rs("M_LEVEL")
if MyRank = 1 then
MyRank = "General User"
elseif MyRank = 2 then
MyRank = "Moderator" 
elseif MyRank = 3 then
MyRank = "Administrator" 
end if
%>
<tr>
<td bgcolor="<% =strForumCellColor %>" align="center">
<% '--------Does user want spam?
if rs("M_RECMAIL") = "1" then
%>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><B>X</B></font>
<% else %>
<input type="checkbox" name="ID" value="<% =rs("MEMBER_ID") %>"><input type="hidden" name="Mail_ALL" value="<% =rs("MEMBER_ID") %>">
<% end if %>
</td>
  <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="pop_profile.asp?mode=display&id=<%=RS("MEMBER_ID")%>"><% =rs("M_NAME") %></a></font></td>
  <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="mailto:<% =rs("M_EMAIL") %>"><% =rs("M_EMAIL") %></a></font></td>
  <td align=right bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% =rs("M_POSTS") %></font></td><td align=right bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% =MyRank %></td><td align=right bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% =ChkDate(rs("M_DATE"))%></font>
</td><td bgcolor="<% =strForumCellColor %>" align=center>
  <a href="pop_profile.asp?mode=Modify&ID=<% =rs("MEMBER_ID") %>&name=<% =rs("M_NAME") %>"><img src="images/icon_pencil.gif" alt="编辑会员" border="0" hspace="0"></a>
  <a href="JavaScript:openWindow('pop_delete.asp?mode=Member&MEMBER_ID=<% =rs("MEMBER_ID") %>')"><img src="images/icon_trashcan.gif" alt="删除会员" border="0" hspace="0"></a>
  </td>
</tr>
<%rs.MoveNext %>
<%loop %>
<% end if %>
 </table>
 </td>
  </tr>
  <tr>
<td align="center"><BR><input type="button" value="选择本页全部会员" onClick="selectAll(this.form,1)">&nbsp;
<input type="submit" name="action" value="将讯息寄给所选择的会员">
&nbsp;<input type="submit" name="action" value="将讯息寄给所有会员">&nbsp;<input type="reset" name="Reset" value="全部清除"></td>
<BR></tr>
</table>

<%
'---------end pagename zip-o
elseif pagename="send" then
'-------------Start sending
%>
<%
My_ID = request.form("id")
Mail_All = request.form("Mail_All")
if My_ID = "*" then
THISESCUELL = "select * from " & strMemberTablePrefix & "MEMBERS where MEMBER_ID in (" & Mail_All & ")"
THISESCUELL = THISESCUELL & " AND " & strMemberTablePrefix & "MEMBERS.M_STATUS = " & 1
THISESCUELL = THISESCUELL & " AND " & strMemberTablePrefix & "MEMBERS.M_EMAIL <> " & "''"
else
THISESCUELL = "select * from " & strMemberTablePrefix & "MEMBERS where MEMBER_ID in (" & My_ID & ")"
THISESCUELL = THISESCUELL & " AND " & strMemberTablePrefix & "MEMBERS.M_STATUS = " & 1
THISESCUELL = THISESCUELL & " AND " & strMemberTablePrefix & "MEMBERS.M_EMAIL <> " & "''"
end if
strSQL= THISESCUELL
set RS=Server.CreateObject("ADODB.Recordset")
RS.Open strSQL, my_Conn, 1, 3
rs.MoveFirst
while not rs.EOF
strRecipientsName = RS.Fields("M_NAME")
strRecipients = RS.Fields("M_email")
strFrom = strSender
strSubject = request.form("SUBJECT")
strMessage = request.form("MESSAGE")
strArchive = request.form("ARCHIVE")
strL_DATE = DateToStr(now())
%>
<!--#INCLUDE FILE="inc_mail.asp" -->
<%
rs.MoveNext
wend 
%>
<font face="<% =strDefaultFontFace %>"><h2>讯息已寄出</h2></font>
<p><center><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">信件已寄给所选收件人。  <a href="admin_emaillist.asp">按这里</a>继续</center></font></p>
<%
if request.form("save") <> "" then
strSubject = replace(request.form("SUBJECT"),"'","''")
strMessage = replace(request.form("MESSAGE"),"'","''")

	set conn = server.createobject ("adodb.connection")
	conn.open My_Conn
	conn.Execute "insert into FORUM_SPAM (SUBJECT, MESSAGE, F_SENT, ARCHIVE) values (" _
		& "'" & strSubject & "', " _
		& "'" & strMessage & "', " _ 
		& "'" & strL_DATE & "', " _		
		& "'" & request.form("ARCHIVE") & "')"

end if
%>

<%
'---------end sending
elseif pagename="compose" then
'-------------Start composing
%>
<% 
myMSG = Request.Form("MSG")
function getFormObject ()
if Request.ServerVariables("REQUEST_METHOD") = "GET" then
set getFormObject=Request.QueryString
else
set getFormObject=Request.Form
end if
end function
%>
<%
set oFormVars=GetFormObject()
if inStr(ucase(oFormVars("Action")),"SELECTED") > 0 then
selected=true
if oFormVars("ID") = "" then
  %>
 <h1>没有选择收件人</h1>
<%
response.end
end if
strSQL="select * from FORUM_MEMBERS where MEMBER_ID in (" & request.form("id") & ")"
strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_STATUS = " & 1
strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_EMAIL <> " & "''"
else
selected=false
strSQL="select * from FORUM_MEMBERS where MEMBER_ID in (" & request.form("Mail_All") & ")"
strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_STATUS = " & 1
strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.M_EMAIL <> " & "''"
end if

set RS=Server.CreateObject("ADODB.Recordset")
RS.Open strSQL, My_Conn, 1, 2

if myMSG <> "" then
strSql2 = "SELECT * FROM " &strMemberTablePrefix & "SPAM WHERE ID =" & myMSG
set rsSP = Server.CreateObject("ADODB.Recordset")
rsSP.open  strSql2, My_Conn, 3
mySUBJECT = Server.HTMLEncode(rsSP("SUBJECT"))
myMESSAGE = rsSP("MESSAGE")
else                       
end if
Mail_All = request.form("Mail_All")
%>
<h2>传送讯息</h2>
<form method="post" action="admin_emaillist.asp">
<input type="hidden" name="pagename" value="send">
<input type="hidden" name="Mail_All" value="<%= Mail_All %>">
<table bordercolor="<% =strTableBorderColor %>" border="0" cellspacing="0" cellpadding="4">
<tr>
<td><font face="arial">标题：</font></td><td><input type="text" name="SUBJECT" size="50" value="<%= mySUBJECT%>"></td>
</tr>
<tr>
<td colspan="2"><font face="arial">信件内容：</font></td>
</tr>
<tr>
<td colspan="2" align="center"><textarea name="MESSAGE" cols="50" rows="10" wrap="PHYSICAL"><%= myMESSAGE %></textarea></td>
</tr>
</table>

<input type="checkbox" name="save" value="1"> <font face="arial">以何种方式储存本讯息？</font> 
<font face="宋体, Arial, Helvetica" size="2"> 
 <select name="ARCHIVE" size="1">
  <option value="0" SELECTED>&nbsp;即时列表</option>
  <option value="1">&nbsp;档案</option>
</select>
 </font>

 &nbsp;<input type="Submit" value="送出">&nbsp;<input type="reset">
<%
if selected then
%>
<h4><i>此讯息将寄给下列用户：</i></h4>
<TABLE width="50%" BORDER=1 bordercolor="<% =strTableBorderColor %>" CELLSPACING=0 align="center" width="100%">
<TR>
<TH BGCOLOR=<% =strHeadCellColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">会员名</FONT></TH>
<TH BGCOLOR=<% =strHeadCellColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">邮件地址</FONT></TH>
</TR>
<%
On Error Resume Next
RS.MoveFirst
do while Not RS.eof
 %>
<TR VALIGN=TOP>
<td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><input type="hidden" name="ID" value="<%=RS("MEMBER_ID")%>"><a href="pop_profile.asp?mode=display&id=<%=RS("MEMBER_ID")%>"><% =rs("M_NAME") %></a></FONT></TD>
<td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% =rs("M_EMAIL") %></FONT></TD>
</TR>
<%
RS.MoveNext
loop%>
</table>
<%
else
%>
<h4><i>信息将发送至全部会员.</i></h4>
<input type="hidden" name="ID" value="*"><%
end if
%></form>
<%
else %>
<h1>Unknown Action</h1>
<p>You have specified an unknown action.  Please try again.</p>
<%
response.end
end if
%>
<%
set rsSP = nothing
set rs = nothing
%>
<!--#INCLUDE file="inc_footer.asp" -->
<% else %>
<%Response.Redirect "admin_login.asp" %>
<% end iF %>
