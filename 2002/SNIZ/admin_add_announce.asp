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
<table border="0" width="100%" align="center">
  <tr>
    <td width="33%" align="left" nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    <img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="default.asp">������̳��ҳ</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="admin_home.asp">��̳��������</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="admin_announce_home.asp">��̳��������</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" height=15 width=15 border="0">&nbsp;��������</font></td>
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
	Date_GetMonthName = "һ��" 
	Case 2
	Date_GetMonthName = "����" 
	Case 3
	Date_GetMonthName = "����" 
	Case 4
	Date_GetMonthName = "����" 
	Case 5
	Date_GetMonthName = "����" 
	Case 6
	Date_GetMonthName = "����" 
	Case 7
	Date_GetMonthName = "����" 
	Case 8
	Date_GetMonthName = "����" 
	Case 9
	Date_GetMonthName = "����" 
	Case 10
	Date_GetMonthName = "ʮ��" 
	Case 11
	Date_GetMonthName = "ʮһ��" 
	Case 12
	Date_GetMonthName = "ʮ����" 
	End Select
	End Function

    	if Request.Form("Method_Type") = "Write_Configuration" then 

	txtSubject = ChkString(Request.Form("Subject"),"title")
	txtMessage = ChkString(Request.Form("Message"),"message")
	txtStartDate = Request.Form("StartYearName") & doublenum(Request.Form("StartMonthName")) & doublenum(Request.Form("StartDayName")) & "000001"
	txtEndDate = Request.Form("EndYearName") & doublenum(Request.Form("EndMonthName")) & doublenum(Request.Form("EndDayName")) & "000001"

	Err_Msg = ""

	if txtSubject = " " then 
		Err_Msg = Err_Msg & "<li>�����������⣡</li>"
	end if

	if txtMessage = " " then 
		Err_Msg = Err_Msg & "<li>��������빫�����ݣ�</li>"
	end if

	if txtStartDate = txtEndDate then 
		Err_Msg = Err_Msg & "<li>���治�ܿ�ʼ�ͽ�����ͬһ�죡</li>"
	end if

	if txtStartDate > txtEndDate then 
		Err_Msg = Err_Msg & "<li>���治���ڿ�ʼǰ�ͽ�����</li>"
	end if

	if Err_Msg = "" then

  		txtMessage = ChkString(Request.Form("Message"),"message")
		txtSubject = ChkString(Request.Form("Subject"),"title")
		strStartDate = Request.Form("StartYearName") & doublenum(Request.Form("StartMonthName")) & doublenum(Request.Form("StartDayName")) & "000001"
		strEndDate = Request.Form("EndYearName") & doublenum(Request.Form("EndMonthName")) & doublenum(Request.Form("EndDayName")) & "000001"

		'## Forum_SQL - Do DB Update
		strSql = "INSERT INTO " & strTablePrefix & "ANNOUNCE ("
		strSql = strSql & "A_AUTHOR"
		strSql = strSql & ", A_SUBJECT"
		strSql = strSql & ", A_MESSAGE"
		strSql = strSql & ", A_START_DATE"
		strSql = strSql & ", A_END_DATE"
		strSql = strSql & ") VALUES ("
		strSql = strSql & "'" & strDBNTUserName & "'"
		strSql = strSql & ", '" & txtSubject & "'"
		strSql = strSql & ", '" & txtMessage & "'"
		strSql = strSql & ", " & "'" & strStartDate & "'"
		strSql = strSql & ", " & "'" & strEndDate & "'"
		strSql = strSql & ")"

		my_Conn.Execute (strSql)

%>
<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">�����Ѿ�������ϣ�</font></p>
<meta http-equiv="Refresh" content="0; URL=admin_announce_home.asp">

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">�޸���ϣ�</font></p>

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="admin_announce_home.asp">�ص���̳��������</font></a></p>
<%		else %>
	<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">������������������û��������ȫ</font></p>

	<table align=center border=0>
	  <tr>
	    <td><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
		<ul>
		<% =Err_Msg %>
		</ul>
	    </font></td>
	  </tr>
	</table>

	<p align=center><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>><a href="JavaScript:history.go(-1)">�뷵����������</a></font></p>
<%		end if %>
<%	else %>
<form action="admin_add_announce.asp" method="post" id="PostTopic" name="PostTopic">
<input type="hidden" name="Method_Type" value="Write_Configuration">
<table border="0" cellspacing="0" cellpadding="0" align=center>
  <tr>
    <td bgcolor="<% =strPopUpBorderColor %>">
<table border="0" cellspacing="1" cellpadding="1">
  <tr valign="top">
    <td align="center" bgcolor="<% =strHeadCellColor %>" colspan="2"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>"><b>��������</b></font></td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">���⣺</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><input maxLength="50" name="Subject" value="<% =Trim(ChkString(TxtSub,"display")) %>" size="30"></td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">UBB����ģʽ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>" align="left"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    <select name="font" onChange="thelp(this.options[this.selectedIndex].value)">
    	<option value="1">����&nbsp;</option>
	<option selected value="2">��ȫ&nbsp;</option>
	<option value="0">����&nbsp;</option>
    </select></font></td>
  </tr>
<!--#INCLUDE FILE="inc_post_buttons.asp" -->  
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�������ݣ�</font><br>
        <br>
        <table border=0>
          <tr>
            <td align="left" nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
<% if strAllowHTML = "1" then %>
            * HTML�﷨����<br>
<% else %>
            * HTML�﷨�ر�<br>
<% end if %>
<% if strAllowForumCode = "1" then %>
            * <a href="JavaScript:openWindow3('pop_forum_code.asp')">����ר�ô���</a>����<br>
<% else %>
            * ����ר�ô���ر�<br>
<% end if %>
            </font>
            </td>
          </tr>
        </table>
    <td bgColor="<% =strPopUpTableColor %>"><textarea name="Message" cols="33" rows=6 wrap="virtual"><% =Trim(cleancode(TxtMsg)) %></textarea></td>
  </tr>
<TR>
<TD align="right" bgcolor="<% =strPopUpTableColor %>" colspan="1">
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">���濪ʼ���ڣ�</font>
</td>
<TD align="left" bgcolor="<% =strPopUpTableColor %>" colspan="1">
<table border="0" cellspacing="1" cellpadding="1">
  <tr valign="top">
<td><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>"><B>��</B></FONT></td>
<td><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>"><B>��</B></FONT></td>
<td><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>"><B>��</B></FONT><br></td>
</tr>
<tr>
<td><SELECT NAME="StartMonthName">
  	<% Date_FillBoxWithMonths "", True %>
    </SELECT></td>
<td><SELECT NAME="StartDayName">
        <% Date_FillBoxWithDays "" %>
    </SELECT></td>
<td><SELECT NAME="StartYearName">
	<% Date_FillBoxWithYears 2000, 2005, "" %>
    </SELECT></td>
</TR>
</TABLE>
</TD>
</TR>
<TR>
<TD align="right" bgcolor="<% =strPopUpTableColor %>" colspan="1">
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">����������ڣ�</font>
</td>
<TD align="left" bgcolor="<% =strPopUpTableColor %>" colspan="1">
<table border="0" cellspacing="1" cellpadding="1">
  <tr valign="top">
<td><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>"><B>��</B></FONT></td>
<td><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>"><B>��</B></FONT></td>
<td><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>"><B>��</B></FONT><br></td>
</tr>
<tr>
<td><SELECT NAME="EndMonthName">
  	<% Date_FillBoxWithMonths DateAdd( "m", +1, Now() ), True %>
    </SELECT></td>
<td><SELECT NAME="EndDayName">
        <% Date_FillBoxWithDays "" %>
    </SELECT></td>
<td><SELECT NAME="EndYearName">
	<% Date_FillBoxWithYears 2000, 2005, Date_GetYear(DateAdd( "m", +1, Now() )) %>
    </SELECT></td>
</TR>
</TABLE>
</TD>
</TR>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" colspan="2" align="center">&nbsp;</td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" colspan="2" align="center"><input type="submit" value="��������" id="submit1" name="submit1"> <input name="Preview" type="button" value="Ԥ������" onclick="OpenPreview()"> <input type="reset" value="�����趨" id="reset1" name="reset1"></td>
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