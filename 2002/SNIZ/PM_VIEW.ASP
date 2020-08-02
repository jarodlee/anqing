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

<% if Request.Cookies(strUniqueID & "User")("Name") = "" Then %>
<meta http-equiv="Refresh" content="2; URL=default.asp">
<center>
<br><br><br><br>
<font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><B>登陆论坛才能使用这个功能</b></font>
<br>
<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="<% =Request.ServerVariables("HTTP_REFERER") %>">返回论坛</font></a></p>
<br><br><br><br><br>
</center>
<% else  %>
<center>

<script language="JavaScript">
<!-- hide from JavaScript-challenged browsers
function pmmembers() { var MainWindow = window.open ("pm_members.asp", "","toolbar=no,location=no,menubar=no,scrollbars=yes,width=350,height=500,top=100,left=100,resizeable=yes,status=no");
}
function ignorelist() { var MainWindow = window.open ("pm_ignore.asp", "","toolbar=no,location=no,menubar=no,scrollbars=yes,width=250,height=500,top=100,left=100,resizeable=yes,status=no");
}
// done hiding -->
</script>
<form action="pm_delete2.asp" method="post" name="DeleteTopic">
<table border="0" width="95%" align="center">
  <tr>
    <td width="33%" align="left" nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    <img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="default.asp">返回论坛首页</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="pm_view.asp">悄悄话收件箱</a>
    </font></td>
  </tr>
</table>

<table border="0" width="95%" cellspacing="1" cellpadding="4">
  <tr>
    <td colspan=2 align=right><% if Request.Cookies("paging")("OUTBOX") = "double" and Request.QueryString("view") <> "outbox" Then %><A HREF="pm_view.asp?view=outbox"><img src="<%=strImageURL %>pmoutbox.jpg" BORDER=0 ALT="发件箱"></A><% end if %><% if Request.Cookies("paging")("OUTBOX") = "double" and Request.QueryString("view") = "outbox" Then %><A HREF="pm_view.asp"><img src="<%=strImageURL %>pminbox.jpg" BORDER=0 ALT="Inbox"></A><% end if %><A HREF="privatesend.asp?method=Topic"><img src="<%=strImageURL %>pmnew.jpg" BORDER=0 ALT="写一条新的悄悄话信息"></A><A HREF="pm_view.asp"><img src="<%=strImageURL %>pmreceive.jpg" BORDER=0 ALT="检查是否有新讯息"></A><A HREF="javascript:ignorelist()"><img src="<%=strImageURL %>pmignore.jpg"  BORDER=0 ALT="查看黑名单列表"></A><A HREF="javascript:pmmembers()"><img src="<%=strImageURL %>pmmemberlist.jpg" BORDER=0 ALT="浏览会员列表"></A><A HREF="javascript:openWindow3('pm_options.asp')"><img src="<%=strImageURL %>pmprivpref.jpg" BORDER=0 ALT="设置你的悄悄话参数"></A></td>
  </tr>
</table>

<% if Request.QueryString("view") <> "outbox" then %>
    <table border="0" width="100%" cellspacing="1" cellpadding="4" align="center">
    <tr>
    <td width=50% >
        <font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="4"><b><% =Request.Cookies(strUniqueID & "User")("Name") %>的悄悄话收件箱</b></font>
    </tr>
    </table>
<table bgcolor="<% =strTableBorderColor %>" cellpadding="4" border="0" cellspacing="1" width="100%">
  <tr bgcolor="<% =strCategoryCellColor %>">
    <td width=20>&nbsp;</td>
    <td width=50%><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>"><B>主题</B></FONT></td>
    <td width=200 nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>"><B>来自</B></FONT></td>
    <td nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>"><B>发送日期</B></FONT></td>
    <td width=8%><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>"><B>删除</B></FONT></td>
  </tr>



<%

     if Request.QueryString("marknew") = "1" then
	     strSql = "UPDATE " & strTablePrefix & "PM "
	     strSql = strSql & " SET " & strTablePrefix & "PM.M_READ = 0 "
	     strSql = strSql & " WHERE (" & strTablePrefix & "PM.M_ID = " & Request.QueryString("id") & ");"
	     my_conn.Execute (strSql)
      end if
      
Response.Flush

	strSql = "SELECT "   & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME,  " & strTablePrefix & "PM.M_ID,  " & strTablePrefix & "PM.M_TO, " & strTablePrefix & "PM.M_SUBJECT, " & strTablePrefix & "PM.M_SENT, " & strTablePrefix & "PM.M_FROM, " & strTablePrefix & "PM.M_READ "
	strSql = strSql & " FROM " & strTablePrefix & "MEMBERS , " & strTablePrefix & "PM "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & Request.Cookies(strUniqueID & "User")("Name") & "'"
	strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "PM.M_TO "
	strSql = strSql & " ORDER BY " & strTablePrefix & "PM.M_SENT DESC" 
	
	Set rsMessage = my_Conn.Execute(strSql)
	



	
	if rsMessage.EOF or rsMessage.BOF then  '## No Private Messages found in DB

%>	
<tr bgcolor="<% =strForumFirstCellColor %>">
  <td>&nbsp;</td>
  <td colspan="4"><font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>你目前没有任何悄悄话讯息</b></font></td>
</tr>
		
<%	else
	
		do Until rsMessage.EOF
		
strSql = "SELECT "   & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME,  " & strTablePrefix & "PM.M_ID  "
strSql = strSql & " FROM " & strTablePrefix & "MEMBERS , " & strTablePrefix & "PM "
strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & rsMessage("M_FROM") & ""

		
Set rsFrom = my_Conn.Execute(strSql)	
	
%>

  <tr bgcolor="<% =strForumFirstCellColor %>">
    <td align=center bgcolor="<% =CColor %>"><b>
    <% if rsMessage("M_READ") = "0" then %>
      <img alt="新讯息" src="<%=strImageURL %>icon_folder_new.gif" width="15" height="15">
<% else %>
      <a href="pm_view.asp?marknew=1&id=<% =rsMessage("M_ID") %>"><img alt="将讯息标示为未阅读" src="<%=strImageURL %>icon_folder.gif" width="15" height="15" border="0"></a>
<% end if %>
    </td>
    <td bgcolor="<% =CColor %>"><font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="privateread.asp?id=<% =rsMessage("M_ID") %>"><% =rsMessage("M_SUBJECT") %></a></font></td>
    
    <td bgcolor="<% =CColor %>"><font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="pop_profile.asp?mode=display&id=<% =rsMessage("M_FROM") %>" target="_new"><% =rsFrom("M_NAME") %></a></font></td>
    <td nowrap bgcolor="<% =CColor %>"><font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% =ChkDate(rsMessage("M_SENT")) %>&nbsp;&nbsp;<% =ChkTime(rsMessage("M_SENT")) %></font></td>
    <td align="center" bgcolor="<% =CColor %>"><input type=checkbox name="DELETE" value="<% =rsMessage("M_ID") %>"></td>
  </tr>


<%	
		rsMessage.MoveNext
	    Loop
	end if
%>

</table>
<table border="0" cellpadding="4" border="0" cellspacing="1" width="100%">
  <tr>
    <td bgcolor="<% =CColor %>" colspan="5" align="right">
      <input type=submit value="删除悄悄话"></form>

    </td>
  </tr>
</table>
<table border="0" cellpadding="4" border="0" cellspacing="1" width="100%">
  <tr>
    <td bgcolor="<% =CColor %>" colspan="5" align="left">
      <p><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>"> 
      <img alt="新讯息" src="<%=strImageURL %>icon_folder_new.gif" width="15" height="15"> 新的悄悄话<br>
      <img alt="旧讯息" src="<%=strImageURL %>icon_folder.gif" width="15" height="15"> 旧的悄悄话
      </font></p>
    </td>
  </tr>
</table>



<% end if 

if Request.Cookies("paging")("OUTBOX") = "single" or Request.QueryString("view") = "outbox" then %>
<center>
<form action="pm_delete2.asp" method="post" name="RemoveTopic">
<input name="RemoveTopic" type="hidden" value="1">
    <table border="0" width="95%" cellspacing="1" cellpadding="4" align="center">
    <tr>
    <td width=50% >
        <font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="4"><b><% =Request.Cookies(strUniqueID & "User")("Name") %>的悄悄话发件箱</b></font>
    </tr>
    </table>
<table border="0" cellpadding="4" border="0" cellspacing="1" width="95%" bgcolor="<% =strTableBorderColor %>">
  <tr bgcolor="<% =strCategoryCellColor %>">
    <td width=20>&nbsp;</td>
    <td width=50%><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>"><B>主题</B></FONT></td>
    <td width=200 nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>"><B>收件人</B></FONT></td>
    <td nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>"><B>发送日期</B></FONT></td>
    <td width=8%><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>"><B>删除</B></FONT></td>
  </tr>
  
<%
	'#PM SQL get private messages
	strSql = "SELECT "   & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME,  " & strTablePrefix & "PM.M_ID,  " & strTablePrefix & "PM.M_TO, " & strTablePrefix & "PM.M_SUBJECT, " & strTablePrefix & "PM.M_SENT, " & strTablePrefix & "PM.M_FROM, " & strTablePrefix & "PM.M_READ, " & strTablePrefix & "PM.M_OUTBOX "
	strSql = strSql & " FROM " & strTablePrefix & "MEMBERS , " & strTablePrefix & "PM "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & Request.Cookies(strUniqueID & "User")("Name") & "'"
	strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "PM.M_FROM "
	strSql = strSql & " AND " & strTablePrefix & "PM.M_OUTBOX = 1" 
	strSql = strSql & " ORDER BY " & strTablePrefix & "PM.M_SENT DESC" 
	
	Set rsFMessage = my_Conn.Execute(strSql)
	
	if rsFMessage.EOF or rsFMessage.BOF then  '## No Private Messages found in DB
%>	
<tr bgcolor="<% =strForumFirstCellColor %>">
  <td>&nbsp;</td>
  <td colspan="4"><font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>你没有任何悄悄话信息在发件箱</b></font></td>
</tr>
		
<%	else
	
		do Until rsFMessage.EOF
		
		'#PM SQL get Message MemberName
		strSql = "SELECT "   & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME,  " & strTablePrefix & "PM.M_ID  "
		strSql = strSql & " FROM " & strTablePrefix & "MEMBERS , " & strTablePrefix & "PM "
		strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & rsFMessage("M_TO") & ""

		Set rsTo = my_Conn.Execute(strSql)
	
%>

  <tr bgcolor="<% =strForumFirstCellColor %>">
    <td align=center bgcolor="<% =CColor %>"><b>
    <% if rsFMessage("M_READ") = "0" then %>
      <img alt="新讯息" src="<%=strImageURL %>icon_folder_new.gif" width="15" height="15">
<% else %>
      <img alt="将讯息标示为未阅读" src="<%=strImageURL %>icon_folder.gif" width="15" height="15" border="0">
<% end if %>
    </td>
    <td bgcolor="<% =CColor %>"><font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="privatesent.asp?id=<% =rsFMessage("M_ID") %>"><% =rsFMessage("M_SUBJECT") %></a></font></td>
    
    <td bgcolor="<% =CColor %>"><font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="pop_profile.asp?mode=display&id=<% =rsFMessage("M_TO") %>" target="_new"><% =rsTo("M_NAME") %></a></font></td>
    <td nowrap bgcolor="<% =CColor %>"><font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% =ChkDate(rsFMessage("M_SENT")) %>&nbsp;&nbsp;<% =ChkTime(rsFMessage("M_SENT")) %></font></td>
    <td align="center" bgcolor="<% =CColor %>"><input type=checkbox name="Remove" value="<% =rsFMessage("M_ID") %>"></td>
  </tr>


<%	
		rsFMessage.MoveNext
	    Loop
	end if
%>

</table>
<table border="0" cellpadding="4" border="0" cellspacing="1" width="95%">
  <tr>
    <td bgcolor="<% =CColor %>" colspan="5" align="right">
      <input type=submit value="删除悄悄话"></form>
    </td>
  </tr>
</table>
<table border="0" cellpadding="4" border="0" cellspacing="1" width="95%">
  <tr>
    <td bgcolor="<% =CColor %>" colspan="5" align="left">
      <p><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>"> 
      <img alt="讯息尚未阅读" src="<%=strImageURL %>icon_folder_new.gif" width="15" height="15"> 新的悄悄话<br>
      <img alt="讯息已阅读" src="<%=strImageURL %>icon_folder.gif" width="15" height="15"> 旧的
      </font></p>
    </td>
  </tr>
</table>

	

</center>
<% end if 
 end if %>
<!--#INCLUDE FILE="inc_footer.asp" -->