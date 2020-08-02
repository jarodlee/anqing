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
<!--#INCLUDE FILE="config.asp" -->
<!--#INCLUDE FILE="inc_functions.asp" -->
<!--#INCLUDE FILE="inc_top.asp" -->
<script language="JavaScript">
function ChangePage(){
	document.PageNum.submit();
}
</script>
<%
Dim	srchUName
Dim	srchFName
Dim	srchLName
Dim	srchState
Dim srchInitial


Function sGetColspan(lIN, lOUT)	
if (strShowModerators = "1") Then lOut = lOut + 1	
If (mlev = "4" or mlev = "3") then lOut = lOut + 1		
If lOut > lIn then
	sGetColspan = lIN
Else
	sGetColspan = lOUT	
End If
End Function

if Request.QueryString("method") <> "" then
	SortMethod = Request.QueryString("method")
end if
	SearchName = Request("M_NAME")
if SearchName = "" then
	SearchName = Request.Form("M_NAME")
end if


	srchUName = Request("UserName")
	srchFName = Request("FirstName")
	srchLName = Request("LastName")
	srchInitial = Request("INITIAL")
	if srchState <> "" then
		srchState = Request("State")
	end if


mypage = request("whichpage")
if mypage = "" then
	mypage = 1
end if


mySQL = request("strSql")
if mySQL = "" then
	mySQL = SQLtemp
end if
'New Search Code
If Request("mode") = "search"  and (srchUNAME = "1" or srchFNAme = "1" or srchLName = "1" or srchState = "1" or srchInitial = "1" ) then 

	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_STATUS, " & strMemberTablePrefix & "MEMBERS.M_NAME, "
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_LEVEL, " & strMemberTablePrefix & "MEMBERS.M_EMAIL, " & strMemberTablePrefix & "MEMBERS.M_COUNTRY, "
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_HOMEPAGE, " & strMemberTablePrefix & "MEMBERS.M_ICQ, " & strMemberTablePrefix & "MEMBERS.M_YAHOO, "
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_AIM, " & strMemberTablePrefix & "MEMBERS.M_TITLE, " & strMemberTablePrefix & "MEMBERS.M_POSTS, "
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_LASTPOSTDATE, " & strMemberTablePrefix & "MEMBERS.M_LASTHEREDATE, " & strMemberTablePrefix & "MEMBERS.M_DATE, "
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_PMRECEIVE, " & strMemberTablePrefix & "MEMBERS.M_STATE "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
if Request.querystring("link") <> "sort" then
	whereSql = " WHERE (" 
	tmpSql = ""
	if srchState <> "1" then
		if srchUName = "1" then
		tmpSql = tmpSql & strMemberTablePrefix & "MEMBERS.M_NAME LIKE '%" & SearchName & "%' OR "
		tmpSql = tmpSql & strMemberTablePrefix & "MEMBERS.M_USERNAME LIKE '%" & SearchName & "%'"
		end if
		if srchFName = "1" then
			if srchUName = "1" then
					tmpSql = tmpSql & " OR "
			end if
			tmpSql = tmpSql & strMemberTablePrefix & "MEMBERS.M_FIRSTNAME LIKE '%" & SearchName & "%'"
		end if
		if srchLName = "1" then
			if srchFName = "1" or srchUName = "1" then 
				tmpSql = tmpSql & " OR "
			end if
			tmpSql = tmpSql & strMemberTablePrefix & "MEMBERS.M_LASTNAME LIKE '%" & SearchName & "%' "
		end if
		if srchInitial = "1" then 
			tmpSQL = strMemberTablePrefix & "MEMBERS.M_NAME LIKE '" & SearchName & "%'"
		end if
	else
		tmpSql = tmpSql & strMemberTablePrefix & "MEMBERS.M_STATE LIKE '%" & SearchName & "%'"
	end if
	whereSql = whereSql & tmpSql &")"
	Session(strCookieURL & "where_Sql") = whereSql
end if	

	if Session(strCookieURL & "where_Sql") <> "" then
		whereSql = Session(strCookieURL & "where_Sql")
	else
		whereSql = ""
	end if
	strSQL = strSql & whereSql
else
	'## Forum_SQL - Get all active topics from last visit
	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_STATUS, " & strMemberTablePrefix & "MEMBERS.M_NAME, "
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_LEVEL, " & strMemberTablePrefix & "MEMBERS.M_EMAIL, " & strMemberTablePrefix & "MEMBERS.M_COUNTRY, "
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_HOMEPAGE, " & strMemberTablePrefix & "MEMBERS.M_ICQ, " & strMemberTablePrefix & "MEMBERS.M_YAHOO, "
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_AIM, " & strMemberTablePrefix & "MEMBERS.M_TITLE, " & strMemberTablePrefix & "MEMBERS.M_POSTS, "
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_LASTPOSTDATE, " & strMemberTablePrefix & "MEMBERS.M_LASTHEREDATE, " & strMemberTablePrefix & "MEMBERS.M_DATE, "
	strSql = strSql & strMemberTablePrefix & "MEMBERS.M_PMRECEIVE "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	if mlev = 4 then
		strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME <> 'n/a' "
	else
		strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_STATUS = " & 1
	end if
end if
select case SortMethod
	case "nameasc"
		strSql = strSql & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_NAME ASC"
	case "namedesc"
		strSql = strSql & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_NAME DESC"
	case "levelasc"
		strSql = strSql & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_TITLE ASC, " & strMemberTablePrefix & "MEMBERS.M_NAME ASC"
	case "leveldesc"
		strSql = strSql & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_TITLE DESC, " & strMemberTablePrefix & "MEMBERS.M_NAME ASC"
	case "lastpostdateasc"
		strSql = strSql & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_LASTPOSTDATE ASC, " & strMemberTablePrefix & "MEMBERS.M_NAME ASC"
	case "lastpostdatedesc"
		strSql = strSql & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_LASTPOSTDATE DESC, " & strMemberTablePrefix & "MEMBERS.M_NAME ASC"
	case "lastheredateasc"
		strSql = strSql & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_LASTHEREDATE ASC, " & strMemberTablePrefix & "MEMBERS.M_NAME ASC"
	case "lastheredatedesc"
		strSql = strSql & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_LASTHEREDATE DESC, " & strMemberTablePrefix & "MEMBERS.M_NAME ASC"
	case "dateasc"
		strSql = strSql & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_DATE ASC, " & strMemberTablePrefix & "MEMBERS.M_NAME ASC"
	case "datedesc"
		strSql = strSql & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_DATE DESC, " & strMemberTablePrefix & "MEMBERS.M_NAME ASC"
	case "countryasc"
		strSql = strSql & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_COUNTRY ASC, " & strMemberTablePrefix & "MEMBERS.M_NAME ASC"
	case "countrydesc"
		strSql = strSql & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_COUNTRY DESC, " & strMemberTablePrefix & "MEMBERS.M_NAME ASC"
	case "postsasc"
		strSql = strSql & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_POSTS ASC, " & strMemberTablePrefix & "MEMBERS.M_NAME ASC"
	case else
		strSql = strSql & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_POSTS DESC, " & strMemberTablePrefix & "MEMBERS.M_NAME ASC"
end select
if strDBType = "mysql" then 'MySql specific code
	if mypage > 1 then 
		OffSet = CInt((mypage - 1) * strPageSize)
		strSql = strSql & " LIMIT " & OffSet & ", " & strPageSize & " "
	end if

	'## Forum_SQL - Get the total pagecount 
	strSql2 = "SELECT COUNT(" & strMemberTablePrefix & "MEMBERS.MEMBER_ID) AS PAGECOUNT "
	strSql2 = strSql2 & " FROM " & strMemberTablePrefix & "MEMBERS " 
	if mlev = 4 then
		strSql2 = strSql2 & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME <> 'n/a' "
	else
		strSql2 = strSql2 & " WHERE " & strMemberTablePrefix & "MEMBERS.M_STATUS = " & 1
	end if


	set rsCount = my_Conn.Execute(strSql2)
	if not rsCount.eof then
		maxpages = (rsCount("PAGECOUNT")  \ strPageSize )
			if rsCount("PAGECOUNT") mod strPageSize <> 0 then
				maxpages = maxpages + 1
			end if
		maxRecs = cint(strPageSize) * maxPages
	else
		maxpages = 0
	end if 

	rsCount.close
	
	set rs = Server.CreateObject("ADODB.Recordset")

	rs.open  strSql, my_Conn, 3
	
	if not (rs.EOF or rs.BOF) then
		rs.movefirst
	end if
 
else 'end MySql specific code

	Set rs = Server.CreateObject("ADODB.RecordSet")
	rs.cachesize=20

	rs.open  strSql, my_conn, 3

	if not (rs.EOF or rs.BOF) then  '## No members found in DB
		rs.movefirst
		rs.pagesize = strPageSize
		rs.cacheSize = strPageSize
		maxPages = cint(rs.pageCount)
		maxRecs = cint(rs.pageSize)
		rs.absolutePage = myPage
		maxpages = cint(rs.pagecount)
	end if
end if
%>
<table width="95%" border="0">
  <tr>
    <td>
    <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><img src="<%=strImageURL %>icon_folder_open.gif" border="0">&nbsp;<a href="default.asp">返回论坛首页</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" border="0">&nbsp;会员列表</font>
    </td>
    <td align="right">
<% if maxpages > 1 then %>
    <table border=0 align="right">
      <tr>
        <td valign="top"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>页数：</b> &nbsp;&nbsp;</font></td>
        <td valign="top"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% Call Paging2() %></font></td>
      </tr>
    </table>
<% else %>
    &nbsp;
<% end if %>
    </td>
  </tr>
</table>
<br>
<table border="0" width="80%" cellspacing="1" cellpadding="4" align="center" bgcolor="<% =strTableBorderColor %>">
 <tr>
 <form action="members.asp?method=<%=SortMethod %>" method="post" name="SearchMembers">
   <td bgcolor="<% =strPopUpTableColor %>"><font colour="<% =strDefaultFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">搜索：&nbsp;  
  <input type="checkbox" name="UserName" value="1" <% if srchUName <> "" then response.write "checked" %> >用户名
  &nbsp;&nbsp;<input type="checkbox" name="FirstName" value="1" <% if srchFName <> "" then response.write "checked" %> >名字
  &nbsp;&nbsp;<input type="checkbox" name="LastName" value="1" <% if srchLName <> "" then response.write "checked" %>>姓氏
  <% if Trim(strState)  = "1" then %>
	&nbsp;&nbsp;<input type="checkbox" name="State" value="1" <% if srchState <> "" then response.write "checked" %>>状态
  <% End If %>
  </font></td>
  <td bgcolor="<% =strPopUpTableColor %>"><font colour="<% =strDefaultFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">查找：&nbsp;<input type="text" name="M_NAME" size="20" value="<%= SearchName %>"></font></td>
<input type="hidden" name="mode" value="search">
<input type="hidden" name="initial" value="0">
  <td bgcolor="<% =strPopUpTableColor %>">
   <INPUT src=<%=strImageURL %>button_go.gif alt="快速搜索" type="image" value="search" id=submit1 name=submit1 border=0>
  </td>
 </tr> 
 </form> 
  <tr bgcolor="<% =strPopUpTableColor %>">
    <td colspan="3" align="center" valign="top"><font colour="<% =strDefaultFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">        按字母排序：
	<a href="members.asp?mode=search&M_NAME=&initial=1&method=<%=SortMethod %> ">全部</a>
	<a href="members.asp?mode=search&M_NAME=A&initial=1&method=<%=SortMethod %> ">A</a>
	<a href="members.asp?mode=search&M_NAME=B&initial=1&method=<%=SortMethod %> "">B</a>
	<a href="members.asp?mode=search&M_NAME=C&initial=1&method=<%=SortMethod %> "">C</a>
	<a href="members.asp?mode=search&M_NAME=D&initial=1&method=<%=SortMethod %> "">D</a>
	<a href="members.asp?mode=search&M_NAME=E&initial=1&method=<%=SortMethod %> "">E</a>
	<a href="members.asp?mode=search&M_NAME=F&initial=1&method=<%=SortMethod %> "">F</a>
	<a href="members.asp?mode=search&M_NAME=G&initial=1&method=<%=SortMethod %> "">G</a>
	<a href="members.asp?mode=search&M_NAME=H&initial=1&method=<%=SortMethod %> "">H</a>
	<a href="members.asp?mode=search&M_NAME=I&initial=1&method=<%=SortMethod %> "">I</a>
	<a href="members.asp?mode=search&M_NAME=J&initial=1&method=<%=SortMethod %> "">J</a>
	<a href="members.asp?mode=search&M_NAME=K&initial=1&method=<%=SortMethod %> "">K</a>
	<a href="members.asp?mode=search&M_NAME=L&initial=1&method=<%=SortMethod %> "">L</a>
	<a href="members.asp?mode=search&M_NAME=M&initial=1&method=<%=SortMethod %> "">M</a>
	<a href="members.asp?mode=search&M_NAME=N&initial=1&method=<%=SortMethod %> "">N</a>
	<a href="members.asp?mode=search&M_NAME=O&initial=1&method=<%=SortMethod %> "">O</a>
	<a href="members.asp?mode=search&M_NAME=P&initial=1&method=<%=SortMethod %> "">P</a>
	<a href="members.asp?mode=search&M_NAME=Q&initial=1&method=<%=SortMethod %> "">Q</a>
	<a href="members.asp?mode=search&M_NAME=R&initial=1&method=<%=SortMethod %> "">R</a>
	<a href="members.asp?mode=search&M_NAME=S&initial=1&method=<%=SortMethod %> "">S</a>
	<a href="members.asp?mode=search&M_NAME=T&initial=1&method=<%=SortMethod %> "">T</a>
	<a href="members.asp?mode=search&M_NAME=U&initial=1&method=<%=SortMethod %> "">U</a>
	<a href="members.asp?mode=search&M_NAME=V&initial=1&method=<%=SortMethod %> "">V</a>
	<a href="members.asp?mode=search&M_NAME=W&initial=1&method=<%=SortMethod %> "">W</a>
	<a href="members.asp?mode=search&M_NAME=X&initial=1&method=<%=SortMethod %> "">X</a>
	<a href="members.asp?mode=search&M_NAME=Y&initial=1&method=<%=SortMethod %> "">Y</a>
	<a href="members.asp?mode=search&M_NAME=Z&initial=1&method=<%=SortMethod %> "">Z</a><br>
	</font></td>
  </tr>
</table>
<br>
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td bgcolor="<% =strTableBorderColor %>">
    <table border="0" width="100%" cellspacing="1" cellpadding="3">
      <tr>
        <td align="center" bgcolor="<% =strHeadCellColor %>">&nbsp;&nbsp;</td>
        <td align="center" bgcolor="<% =strHeadCellColor %>"><a href="members.asp?link=sort&mode=search&M_NAME=<%=SearchName %>&method=<% if Request.QueryString("method") = "nameasc" then Response.Write("namedesc") else Response.Write("nameasc") end if %>"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">会员名称</font></b></a></td>
        <td align="center" bgcolor="<% =strHeadCellColor %>"><a href="members.asp?link=sort&mode=search&M_NAME=<%=SearchName %>&method=<% if Request.QueryString("method") = "levelasc" then Response.Write("leveldesc") else Response.Write("levelasc") end if %>"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">状态</font></b></a></td>
        <td align="center" bgcolor="<% =strHeadCellColor %>"><a href="members.asp?link=sort&mode=search&M_NAME=<%=SearchName %>&method=<% if Request.QueryString("method") = "postsdesc" then Response.Write("postsasc") else Response.Write("postsdesc") end if %>"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">文章数</font></b></a></td>
        <td align="center" bgcolor="<% =strHeadCellColor %>"><a href="members.asp?link=sort&mode=search&M_NAME=<%=SearchName %>&method=<% if Request.QueryString("method") = "lastpostdatedesc" then Response.Write("lastpostdatedasc") else Response.Write("lastpostdatedesc") end if %>"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">最后发表</font></b></a></td>
        <td align="center" bgcolor="<% =strHeadCellColor %>"><a href="members.asp?link=sort&mode=search&M_NAME=<%=SearchName %>&method=<% if Request.QueryString("method") = "datedesc" then Response.Write("dateasc") else Response.Write("datedesc") end if %>"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">注册日期</font></b></a></td>
        <td align="center" bgcolor="<% =strHeadCellColor %>"><a href="members.asp?link=sort&mode=search&M_NAME=<%=SearchName %>&method=<% if Request.QueryString("method") = "countryasc" then Response.Write("countrydesc") else Response.Write("countryasc") end if %>"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">国籍</font></b></a></td>
<% if mlev = 4 or mlev = 3 then %>
        <td align="center" bgcolor="<% =strHeadCellColor %>"><a href="members.asp?method=<% if Request.QueryString("method") = "lastheredatedesc" then Response.Write("lastheredatedasc") else Response.Write("lastheredatedesc") end if %>"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">最后登陆时间</font></b></a></td>
<% end if %>
<% if mlev = 4 then %>
        <td align="center" bgcolor="<% =strHeadCellColor %>"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">&nbsp;</font></b></td>
<% end if %>
      </tr>
<% if rs.EOF or rs.BOF then  '## No Members Found in DB %>
      <tr>
        <td colspan="<%=sGetColspan(9, 7)%>" bgcolor="<% =strForumCellColor %>" ><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>目前没有任何会员</b></font></td>
      </tr>
<% else %>
<%	currMember = 0 %>
<%
	i = 0
	howManyRecs = 0
	rec = 1
	do until rs.Eof or rec = (strPageSize + 1)
		if i = 1 then 
			CColor = strAltForumCellColor
		else
			CColor = strForumCellColor
		end if
%>
      <tr>
        <td bgcolor="<% =CColor %>" align=left>
        <% if strUseExtendedProfile then %>
			<a href="pop_profile.asp?mode=display&id=<% =rs("MEMBER_ID") %>">
        <% else %>
			<a href="JavaScript:openWindow3('pop_profile.asp?mode=display&id=<% =rs("MEMBER_ID") %>')">
		<% end if %>	
        <% if rs("M_STATUS") = 0 then 
        %><img src="<%=strImageURL %>icon_profile_locked.gif" alt="查看会员资料" height=15 width=15 border=0 hspace="0"><%
		else 
        %><img src="<%=strImageURL %>icon_profile.gif" alt="查看会员资料" height=15 width=15 border=0 hspace="0"><%
		end if 
        %></a>
<%		if (rs("M_PMRECEIVE") = "1") then %>
	  		<a href="privatesend.asp?method=Topic&mname=<% =rs("M_NAME") %>"><img src="<%=strImageURL %>pm.gif" width="11" height="17" alt="发送悄悄话讯息" border="0" hspace="0"></a>
<%		end if  %>
<%		if strICQ = "1" and Trim(rs("M_ICQ")) <> "" then %>
        <a href="JavaScript:openWindow('pop_messengers.asp?mode=ICQ&ICQ=<% =ChkString(rs("M_ICQ"), "JSurlpath") %>&M_NAME=<% =ChkString(rs("M_NAME"),"display") %>')"><img src="http://online.mirabilis.com/scripts/online.dll?icq=<% =chkstring(rs("M_ICQ"), "urlpath") %>&img=5" height=15 width=15 alt="发送ICQ讯息给 <% =ChkString(rs("M_NAME"),"display") %>" border="0" align="absmiddle" hspace="0"></a>
<%		end if %>
<%		if strYAHOO = "1" and Trim(rs("M_YAHOO")) <> "" then %>
<a href="http://search.tencent.com/cgi-bin/friend/user_show_info?ln=<% =rs("M_YAHOO") %>"  target=_blank><img align=absmiddle width=16 height=16 src="http://icon.tencent.com/<% =rs("M_YAHOO") %>/s/00/" alt="<% =rs("M_YAHOO") %>" border=0></a>
<%		end if %>
<%		if strAIM = "1" and Trim(rs("M_AIM")) <> "" then %>
        <a href="JavaScript:openWindow('pop_messengers.asp?mode=AIM&AIM=<% =ChkString(rs("M_AIM"), "JSurlpath") %>&M_NAME=<% =ChkString(rs("M_NAME"),"JSurlpath") %>')"><img src="<%=strImageURL %>icon_aim.gif" height=15 width=15 alt="发送aim给<% =ChkString(rs("M_NAME"),"display") %>" border="0" align="absmiddle" hspace="0"></a>
<%		end if %>
        </td>
        <td bgcolor="<% =CColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
        <% if strUseExtendedProfile then %>
			<a href="pop_profile.asp?mode=display&id=<% =rs("MEMBER_ID") %>">
        <% else %>
			<a href="JavaScript:openWindow3('pop_profile.asp?mode=display&id=<% =rs("MEMBER_ID") %>')">
		<% end if %>	
		<% =ChkString(rs("M_NAME"),"display") %></a></font></td>
        <td bgcolor="<% =CColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% =ChkString(getMember_Level(rs("M_TITLE"), rs("M_LEVEL"), rs("M_POSTS")),"display") %></font></td>
        <td bgcolor="<% =CColor %>" align=center><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strForumFontColor %>">
<%		if IsNull(rs("M_POSTS")) then %>
        -
<%		else %>
         <% =rs("M_POSTS") %>
<%			if strShowRank = 2 or strShowRank = 3 then 
%>
        <br><% Response.write(getStar_Level(rs("M_LEVEL"), rs("M_POSTS"))) %>
<%			end if %>
<%		end if %>
        </font></td>
        <% if IsNull(rs("M_LASTPOSTDATE")) or Trim(rs("M_LASTPOSTDATE")) = "" then%>
        <td bgcolor="<% =CColor %>" align=center nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strForumFontColor %>">-</font></td>
        <% else %>
        <td bgcolor="<% =CColor %>" align=center nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strForumFontColor %>"><% =ChkDate(rs("M_LASTPOSTDATE")) %></font></td>
        <% end if %>
        
        <td bgcolor="<% =CColor %>" align=center nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strForumFontColor %>"><% =ChkDate(rs("M_DATE")) %></font></td>
        <td bgcolor="<% =CColor %>" align=center><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% =rs("M_COUNTRY") %>&nbsp;</font></td>
<%		if mlev = 4 or mlev = 3 then %>
        <td bgcolor="<% =CColor %>" align=center nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strForumFontColor %>"><% =ChkDate(rs("M_LASTHEREDATE")) %></font></td>
<%		end if %>
<%		if mlev = 4 then %>
        <td bgcolor="<% =CColor %>" align=center><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
<%			if rs("MEMBER_ID") <> 1 then %>
<%				if rs("M_STATUS") <> 0 then %>
          <a href="JavaScript:openWindow('pop_lock.asp?mode=Member&MEMBER_ID=<% =rs("MEMBER_ID") %>')"><img src="<%=strImageURL %>icon_lock.gif" alt="停用此帐号" border="0" hspace="0"></a>
<%				else %>
          <a href="JavaScript:openWindow('pop_open.asp?mode=Member&MEMBER_ID=<% =rs("MEMBER_ID") %>')"><img src="<%=strImageURL %>icon_unlock.gif" alt="开启此帐号" border="0" hspace="0"></a>
<%				end if %>
<%			end if 
			if strUseExtendedProfile then %>
			   <a href="pop_profile.asp?mode=Modify&ID=<% =rs("MEMBER_ID") %>&name=<% =ChkString(rs("M_NAME"),"urlpath") %>"><img src="<%=strImageURL %>icon_pencil.gif" alt="编辑会员" border="0" hspace="0"></a>
<%			else %>
			  <a href="JavaScript:openWindow3('pop_profile.asp?mode=Modify&ID=<% =rs("MEMBER_ID") %>&name=<% =ChkString(rs("M_NAME"),"urlpath") %>')"><img src="<%=strImageURL %>icon_pencil.gif" alt="编辑会员" border="0" hspace="0"></a>
<%			end if
			if rs("MEMBER_ID") = 1 then %>
<%				'## Do Nothing %>
<%			else %>
          <a href="JavaScript:openWindow('pop_delete.asp?mode=Member&MEMBER_ID=<% =rs("MEMBER_ID") %>')"><img src="<%=strImageURL %>icon_trashcan.gif" alt="删除会员" border="0" hspace="0"></a>
<%			end if 
			if rs("M_LEVEL") = 1 then %>
          <a href="pop_profile.asp?mode=Moderator&action=add&ID=<% =rs("MEMBER_ID") %>"><img src="<%=strImageURL %>icon_mod.gif" alt="设为版主" border="0" hspace="0"></a>
		<% Elseif  rs("M_LEVEL") = 2 then%>
		  <a href="pop_profile.asp?mode=Moderator&action=del&ID=<% =rs("MEMBER_ID") %>"><img src="<%=strImageURL %>icon_delmod.gif" alt="卸任版主" border="0" hspace="0"></a>			
		  <% End If %>
        </font></b></td>
<%		end if %>
      </tr>
<%
		currMember = rs("MEMBER_ID")
		rs.MoveNext
		i = i + 1
		if i = 2 then i = 0
		rec = rec + 1
	loop 
end if 
%>
    </table>
	
    </td>
  </tr>
  <tr>
    <td colspan="2">
    <table border="0" width="100%">
      <tr>
        <td>
<% if maxpages > 1 then %>
        <table border=0>
          <tr>
            <td valign="top"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>会员列表共有 <% =maxpages %> 页：</b> &nbsp;&nbsp;</font></td>
            <td valign="top"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% Call Paging() %></font></td>
          </tr>
        </table>
<% else %>
        &nbsp;
<% end if %>
        </td>
      </tr>
    </table>
    </td>
  </tr>
</table>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
sub Paging2()
	if maxpages > 1 then
		if Request.QueryString("whichpage") = "" then
			sPageNumber = 1
		else
			sPageNumber = Request.QueryString("whichpage")
		end if
		if Request.QueryString("method") = "" then
			sMethod = "postsdesc"
		else
			sMethod = Request.QueryString("method")
		end if

		sScriptName = Request.ServerVariables("script_name")
		Response.Write("<form name=""PageNum"" action=""members.asp?method=" & sMethod & """>")
		Response.Write("<select name=""whichpage"" size=""1"" onchange=""ChangePage()"">")
		for counter = 1 to maxpages
			if counter <> cint(sPageNumber) then   
				Response.Write "<OPTION VALUE=""" & counter &  """>" & counter
			else
				Response.Write "<OPTION SELECTED VALUE=""" & counter &  """>" & counter
			end if
		next
		Response.Write("</select>")

	end if
end sub 

sub Paging()
	if maxpages > 1 then
		if Request.QueryString("whichpage") = "" then
			sPageNumber = 1
		else
			sPageNumber = Request.QueryString("whichpage")
		end if
		if Request.QueryString("method") = "" then
			sMethod = "postsdesc"
		else
			sMethod = Request.QueryString("method")
		end if

		sScriptName = Request.ServerVariables("script_name")
		Response.Write("<table border=0 width=100% cellspacing=0 cellpadding=1 align=top><tr>")
		for counter = 1 to maxpages
			if counter <> cint(sPageNumber) then   
				sNum = "<td align=right bgcolor=" & strPageBGColor & "><font face=" & strDefaultFontFace & " size=" & strDefaultFontSize & ">" & "&nbsp;" & widenum(counter) &  "<a href=""" & sScriptName 
				sNum = sNum & "?whichpage=" & counter
				'sNum = sNum & "&pagesize=" & mypagesize
				sNum = sNum & "&method=" & sMethod
				sNum = sNum & """>" & counter & "</a></font></td>"
				Response.Write sNum
			else
				Response.Write("<td align=right bgcolor=" & strPageBGColor & "><font face=" & strDefaultFontFace & " size=" & strDefaultFontSize & ">" & "&nbsp;" & widenum(counter) & "<b>" & counter & "</b></font></td>")
			end if
			if counter mod strPageNumberSize = 0 then
				Response.Write("</tr><tr>")
			end if
		next
		Response.Write("</tr></table>")
	end if
end sub 
%>


