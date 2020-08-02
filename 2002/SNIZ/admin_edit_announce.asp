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
<% If Session(strCookieURL & "Approval") = "15916941253" Then %>
<!--#INCLUDE file="inc_functions.asp" -->
<!--#INCLUDE FILE="inc_code.asp" -->
<!--#INCLUDE file="inc_top.asp" -->
<script language="JavaScript">
<!--
function OpenPreview()
{
	var curCookie = "strMessagePreview=" + escape(document.PostTopic.Message.value);
	document.cookie = curCookie;
	popupWin = window.open('pop_preview.asp', 'preview_page', 'scrollbars=yes,width=550,height=250')
}
//-->
</script>
<table border="0" width="95%" align="center">
  <tr>
    <td width="33%" align="left" nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    <img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="default.asp">返回论坛首页</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="admin_home.asp">论坛管理中心</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="admin_announce_home.asp">论坛公告中心</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="admin_review_announce.asp">查看修改公告</a>
<%  if Request.Form("Method_Type") = "Write_Configuration" then %>
    	&nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" height=15 width=15 border="0">&nbsp;<% =ChkString(Request.Form("Subject"),"title") %></font></td>
<%  else %>
    	&nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" height=15 width=15 border="0">&nbsp;<% =ChkString(Request.QueryString("A_SUBJECT"),"display") %></font></td>
<%  end if %>
  </tr>
</table>
<%	Function Date_GetYear( dDate )
	Date_GetYear = DatePart( "yyyy", dDate )
	End Function

	Function Date_GetMonthNo( dDate )
	Date_GetMonthNo = DatePart( "m", dDate )
	End Function

	Function Date_GetDayNo( dDate )
	Date_GetDayNo = DatePart( "d", dDate )
	End Function

	Function Date_GetMonthName( nMonthNo )
	Select Case nMonthNo
	Case 1
	Date_GetMonthName = "一月" 
	Case 2
	Date_GetMonthName = "二月" 
	Case 3
	Date_GetMonthName = "三月" 
	Case 4
	Date_GetMonthName = "四月" 
	Case 5
	Date_GetMonthName = "五月" 
	Case 6
	Date_GetMonthName = "六月" 
	Case 7
	Date_GetMonthName = "七月" 
	Case 8
	Date_GetMonthName = "八月" 
	Case 9
	Date_GetMonthName = "九月" 
	Case 10
	Date_GetMonthName = "十月" 
	Case 11
	Date_GetMonthName = "十一月" 
	Case 12
	Date_GetMonthName = "十二月"
	End Select
	End Function


    	if Request.Form("Method_Type") = "Write_Configuration" then
	txtSubject = ChkString(Request.Form("Subject"),"title")
	txtMessage = ChkString(Request.Form("Message"),"message")
	txtEndDate = Request.Form("EndYearName") & doublenum(Request.Form("EndMonthName")) & doublenum(Request.Form("EndDayName")) & "000001"

	Err_Msg = ""

	if txtSubject = " " then
		Err_Msg = Err_Msg & "<li>你必须输入标题！</li>"
	end if

	if txtMessage = " " then
		Err_Msg = Err_Msg & "<li>你必须输入公告内容！</li>"
	end if

	if txtEndDate < DateToStr(Now()) then
		Err_Msg = Err_Msg & "<li>公告不能在开始前就结束!</li>"
	end if

	if Err_Msg = "" then

  		txtMessage = ChkString(Request.Form("Message"),"message")
		txtSubject = ChkString(Request.Form("Subject"),"title")
		strEndDate = Request.Form("EndYearName") & doublenum(Request.Form("EndMonthName")) & doublenum(Request.Form("EndDayName")) & "000001"

		'## Forum_SQL - Do DB Update
		strSql = "UPDATE " & strTablePrefix & "ANNOUNCE "
		strSql = strSql & " SET A_SUBJECT = '" & txtSubject & "'"
		strSql = strSql & ",    A_MESSAGE = '" & txtMessage & "'"
		strSql = strSql & ",    A_END_DATE = '" & strEndDate & "'"
		strSql = strSql & " WHERE A_ID = " & Request.Form("A_ID")

		my_Conn.Execute (strSql)

%>
<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">公告已经修改！</font></p>
<meta http-equiv="Refresh" content="0; URL=admin_review_announce.asp">

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">修改完成！</font></p>

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="admin_review_announce.asp">返回查看修改公告</font></a></p>
<%		else %>
	<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">你输入的资料有问题或没有输入完全</font></p>

	<table align=center border=0>
	  <tr>
	    <td><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
		<ul>
		<% =Err_Msg %>
		</ul>
	    </font></td>
	  </tr>
	</table>

	<p align=center><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>><a href="JavaScript:history.go(-1)">请返回重新输入</a></font></p>
<%		end if %>
<%	else %>
<%
	'## Forum_SQL - Get Announcements from DB
	strSql = "SELECT * "
	strSql = strSql & " FROM " & strTablePrefix & "ANNOUNCE "
	strSql = strSql & " WHERE A_ID = " & Request.QueryString("A_ID")

	set rs = my_Conn.Execute (strSql)

        TxtSub = rs("A_SUBJECT")
        TxtMsg = rs("A_MESSAGE")
	strStartDate = StrToDate(rs("A_START_DATE"))
	strEndDate = StrToDate(rs("A_END_DATE"))
%>
<form action="admin_edit_announce.asp" method="post" id="PostTopic" name="PostTopic">
<input type="hidden" name="Method_Type" value="Write_Configuration">
<input type="hidden" name="A_ID" value="<% =rs("A_ID") %>">
<table border="0" cellspacing="0" cellpadding="0" align=center>
  <tr>
    <td bgcolor="<% =strPopUpBorderColor %>">
<table border="0" cellspacing="1" cellpadding="1">
  <tr valign="top">
    <td align="center" bgcolor="<% =strHeadCellColor %>" colspan="2"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>"><b>查看修改公告</b></font></td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">标题：</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><input maxLength="50" name="Subject" value="<% =Trim(ChkString(TxtSub,"display")) %>" size="30"></td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">模式：</font></td>
    <td bgColor="<% =strPopUpTableColor %>" align="left"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    <select name="font" onChange="thelp(this.options[this.selectedIndex].value)">
    	<option value="1">帮助&nbsp;</option>
	<option selected value="2">完全&nbsp;</option>
	<option value="0">基本&nbsp;</option>
    </select></font></td>
  </tr>
<!--#INCLUDE FILE="inc_post_buttons.asp" -->
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">公告内容：</font><br>
        <br>
        <table border=0>
          <tr>
            <td align="left" nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
<% if strAllowHTML = "1" then %>
            * HTML语法开启<br>
<% else %>
            * HTML语法关闭<br>
<% end if %>
<% if strAllowForumCode = "1" then %>
            * <a href="JavaScript:openWindow3('pop_forum_code.asp')">本版专用代码</a>开启<br>
<% else %>
            * 本版专用代码关闭<br>
<% end if %>
            </font>
            </td>
          </tr>
        </table>
    <td bgColor="<% =strPopUpTableColor %>"><textarea name="Message" cols="33" rows=6 wrap="virtual"><% =Trim(cleancode(TxtMsg)) %></textarea></td>
  </tr>
<TR>
<TD align="right" bgcolor="<% =strPopUpTableColor %>" colspan="1">
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">公告开始日期：</font>
</td>
<TD align="left" bgcolor="<% =strPopUpTableColor %>" colspan="1">
<table border="0" cellspacing="1" cellpadding="1">
  <tr valign="top">
<td><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>"><B>月</B></FONT></td>
<td><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>"><B>日</B></FONT></td>
<td><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>"><B>年</B></FONT><br></td>
</tr>
<tr>
<td><SELECT NAME="StartMonthName" disabled>
  	<% Date_FillBoxWithMonths DatePart("m", strStartDate), True %>
    </SELECT></td>
<td><SELECT NAME="StartDayName" disabled>
        <% Date_FillBoxWithDays DatePart("d", strStartDate) %>
    </SELECT></td>
<td><SELECT NAME="StartYearName" disabled>
	<% Date_FillBoxWithYears 2000, 2005, DatePart("yyyy", strStartDate) %>
    </SELECT></td>
</TR>
</TABLE>
</TD>
</TR>
<TR>
<TD align="right" bgcolor="<% =strPopUpTableColor %>" colspan="1">
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">公告结束日期：</font>
</td>
<TD align="left" bgcolor="<% =strPopUpTableColor %>" colspan="1">
<table border="0" cellspacing="1" cellpadding="1">
  <tr valign="top">
<td><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>"><B>月</B></FONT></td>
<td><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>"><B>日</B></FONT></td>
<td><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>"><B>年</B></FONT><br></td>
</tr>
<tr>
<td><SELECT NAME="EndMonthName">
  	<% Date_FillBoxWithMonths DatePart("m", strEndDate), True %>
    </SELECT></td>
<td><SELECT NAME="EndDayName">
        <% Date_FillBoxWithDays DatePart("d", strEndDate) %>
    </SELECT></td>
<td><SELECT NAME="EndYearName">
	<% Date_FillBoxWithYears 2000, 2005, DatePart("yyyy", strEndDate) %>
    </SELECT></td>
</TR>
</TABLE>
</TD>
</TR>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" colspan="2" align="center">&nbsp;</td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" colspan="2" align="center"><input type="submit" value="提交修改" id="submit1" name="submit1"> <input name="Preview" type="button" value="预览公告" onclick="OpenPreview()"> <input type="reset" value="重新设定" id="reset1" name="reset1"></td>
  </tr>
</TABLE>
</table>
    </td>
  </tr>
</table>
</form>
</font>
<%	end if %>
<!--#INCLUDE file="inc_footer.asp" -->
<% else %>
<% Response.Redirect "admin_login.asp" %>
<% end if %>
<%
Sub Date_FillBoxWithYears( nFirst, nLast, nDefault )
Dim nYear
If nDefault = "" Then
	nDefault = Date_GetYear( Now() ) 
End If 
For nYear = nFirst To nLast
			Response.Write "<option"
			If nYear = nDefault Then
			Response.Write " selected=yes"
			End If
			Response.Write ">" & nYear & "</option>" & vbcrlf
Next
End Sub

Sub Date_FillBoxWithMonths( nDefault, fUseNumbersAsValue )
Dim nMonth
If nDefault = "" Then
	nDefault = Date_GetMonthNo( Now() )
End If 
For nMonth = 1 To 12
			Response.Write "<option"
			If nMonth = nDefault Then
			Response.Write " selected=yes"
			End If
			If fUseNumbersAsValue = True Then
			Response.Write " value=" & nMonth
			End If
			Response.Write ">" & Date_GetMonthName(nMonth) & "</option>" & vbcrlf
Next
End Sub

Sub Date_FillBoxWithDays( nDefault )
Dim nDay
If nDefault = "" Then
	nDefault = Date_GetDayNo( Now() )
End If
For nDay = 1 To 31
			Response.Write "<option"
			If nDay = nDefault Then
			Response.Write " selected=yes"
			End If
			Response.Write ">" & nDay & "</option>" & vbcrlf
Next
End Sub
%>