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
<%
	' Constants for the days of the week
	Const cSUN = 1, cMON = 2, cTUE = 3, cWED = 4, cTHU = 5, cFRI = 6, cSAT = 7
	
	dim intPrivateEvent
	
	if Request.Form("isPrivateEvent") = "1" then
		intPrivateEvent = 1
	else
		isPrivateEvent = 0
	end if

	' Get the name of this file
	dim sScript
	sScript = Request.ServerVariables("SCRIPT_NAME")
	
	'set the date to today
	dim datToday
	datToday = date()
	
	' Check for valid month input
	dim intThisMonth
	intThisMonth = Request.QueryString("month")
	if intThisMonth = "" then
		intThisMonth = month(datToday)
	else
		intThisMonth = cint(intThisMonth)
	end if
	If intThisMonth < 1 OR intThisMonth > 12 Then
		intThisMonth = Month(datToday)
	End If
	
	' Check for valid year input
	dim intThisYear
	intThisYear = Request.QueryString("year")
	If intThisYear = "" Then
		intThisYear = Year(datToday)
	else
		intThisYear = cint(intThisYear)
	End If

	dim sSQL


	dim sStart_Date, sEnd_Date, sEvent_title, sEvent_Details

	dim sMode, berror
	sMode = Request.QueryString("mode")

	select case sMode
		case "delete" 
			if mlev >= 3 then
				sSQL = "DELETE * FROM FORUM_EVENTS WHERE Event_ID=" & Request.QueryString("Event_ID")
				my_Conn.Execute ssql
				strCMessage = "The selected event was deleted"
			end if
		case "edit"
			if mlev >= 3 then
				sMode = "edit"
				strCMessage = "Event Edit Successful"
			end if
		case "add"
			if mlev >= 1 then
				smode = "add"
			end if
	end select

	dim ddate
	dDate = Request.QueryString("Date")
	If not IsEmpty(ddate) and IsDate(ddate) and smode = "" Then
		sMode = "display"
	end if

	if Request.Form("EVENT") <> "" then
		Update_Event(Request.Form("EVENT"))
	end if	

	dim strMonthName, datFirstDay, intFirstWeekday, intLastDay, intPrevMonth, intNextMonth, intPrevYear, intNextYear
	dim LastMonthDate, NextMonthDate, intPrintDay, intLastMonth, dToday, dFirstDay, dLastDay, endrows, intLoopDay
	dim bevents, i, sTitle
%>
<style TYPE="text/css">
  TD.NON {font-family :Tahoma, 宋体, Arial; font-size :12px; color :#C0C0C0; font-weight :normal;}
  TD.TOP {font-family :Tahoma, 宋体, Arial; font-size :12px; color :#0000FF; font-weight :bold;}
  TD.Some {font-family :Tahoma, 宋体, Arial; font-size :12px; color :#000000; font-weight :normal;}
  TD.Today {font-family :Tahoma, 宋体, Arial; font-size :12px; background-color :khaki; font-weight :normal;}
  TD.Selected {font-family :Tahoma, 宋体, Arial; font-size :12px; background-color :silver; font-weight :normal;}
  TD.HL {font-family :Tahoma, 宋体, Arial; font-size :12px; color :#000000; font-weight :normal;}
  TD.NON2 {font-family: Tahoma, 宋体, Arial; font-size :12px; background-color :lightgrey; font-weight :normal}
A.NOEVENT:link
{
    COLOR: #000000;
    CURSOR: nw-resize;
    FONT-FAMILY: Tahoma, 宋体, Arial;
    FONT-WEIGHT: normal;
    TEXT-DECORATION: none
}
A.NOEVENT:visited
{
    COLOR: #000000;
    CURSOR: nw-resize;
    FONT-FAMILY: Tahoma, 宋体, Arial;
    FONT-WEIGHT: normal;
    TEXT-DECORATION: none
}
A.NOEVENT:hover
{
    BACKGROUND-COLOR: silver;
    COLOR: #000000;
    CURSOR: nw-resize;
    FONT-FAMILY: Tahoma, 宋体, Arial;
    FONT-WEIGHT: normal;
    TEXT-DECORATION: none
}
A.EVENT:link
{
    COLOR: green;
    CURSOR: nw-resize;
    FONT-FAMILY: Tahoma, 宋体, Arial;
    FONT-WEIGHT: bold;
    TEXT-DECORATION: none
}
A.EVENT:visited
{
    COLOR: green;
    CURSOR: nw-resize;
    FONT-FAMILY: Tahoma, 宋体, Arial;
    FONT-WEIGHT: bold;
    TEXT-DECORATION: none
}
A.EVENT:hover
{
    BACKGROUND-COLOR: khaki;
    COLOR: green;
    CURSOR: nw-resize;
    FONT-FAMILY: Tahoma, 宋体, Arial;
    FONT-WEIGHT: bold;
    TEXT-DECORATION: none
}
A.MONTH:link
{
    COLOR: black;
    FONT-FAMILY: Tahoma, 宋体, Arial;
    FONT-WEIGHT: normal;
    TEXT-DECORATION: none
}
A.MONTH:visited
{
    COLOR: black;
    FONT-FAMILY: Tahoma, 宋体, Arial;
    FONT-WEIGHT: normal;
    TEXT-DECORATION: none
}
A.MONTH:hover
{
    COLOR: black;
    FONT-FAMILY: Tahoma, 宋体, Arial;
    FONT-WEIGHT: normal;
    TEXT-DECORATION: underline
}
A.NORMAL:link
{
    COLOR: #000000;
    FONT-FAMILY: 宋体, Arial;
    FONT-WEIGHT: normal;
    TEXT-DECORATION: none
}
A.NORMAL:visited
{
    COLOR: #000000;
    FONT-FAMILY: 宋体, Arial;
    FONT-WEIGHT: normal;
    TEXT-DECORATION: none
}
A.NORMAL:hover
{
    COLOR: #000000;
    FONT-FAMILY: 宋体, Arial;
    FONT-WEIGHT: normal;
    TEXT-DECORATION: none
}
.createdby
{
    FONT-STYLE: italic;
    TEXT-ALIGN: right
}
.bluetext
{
    COLOR: #336699
}
.month
{
    FONT-WEIGHT: bold
}
.eventname
{
    FONT-WEIGHT: bold
}
</style>
<table bgcolor="<% =strTableBorderColor %>" border="0" cellpadding="0" cellspacing="1" width="100%" ALIGN=CENTER>
	<tr>
		<td bgcolor="<% =strForumCellColor %>" align="center" valign="top" nowrap width="140">
			<table border="0" cellpadding="1" cellspacing="2" width="140" height="100%">
				<tr>
					<td align="center" bgcolor="<% =strHeadCellColor %>"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">显示月份</font></b></td>
				</tr>
				<%emitmonths()%>
				<tr>
					<td align=center>
						<FORM ACTION="<% =sScript%>" METHOD=GET id=frmSelectMonth name=frmSelectMonth>
							<SELECT NAME=MONTH>
								<%for i = 1 to 12%>
								<OPTION VALUE=<%=i%> <%if i = intThisMonth then Response.Write(" selected")%>><%= monthname(i,true)%>
								<%next%>
							</SELECT>
							<SELECT NAME=YEAR>
								<%for i = -3 to 3%>
								<OPTION VALUE=<%= intThisYear + i%> <%if (intThisYear + i) = intThisYear then Response.Write(" selected")%>><%= intThisYear + i%>
								<%next%>
							</SELECT>
							<INPUT TYPE="submit" NAME="Go" value="显示" BORDER="0" WIDTH="35" HEIGHT="20">
						</FORM>
					</td>
				</tr>
				<%if mlev >= 1 then%>
				<tr>
					<td align=center>
						<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a HREF="events.asp?mode=add&date=<%if Request.QueryString("Date")= "" then Response.Write(date()) else Response.Write(Request.QueryString("Date"))%>">新增事项</A></font>
					</td>
				</tr>
				<%end if%>
				<%if smode <> "" then%>
				<tr>
					<td align=center>
						<div align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href=events.asp>回到月历首页</a></font></div>
					</td>
				</tr>
				<%end if%>
				<tr>
					<td>
					&nbsp;
					</td>
				</tr>
				<tr>
					<td align="center" bgcolor="<% =strHeadCellColor %>"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">
					当前事项：
					</font></b></td>
				</tr>
				<tr>
					<td>
					<% call emitupcomingevents %>
					<br>
					<br>
					</td>
				</tr>
				<tr>
					<td align="center" bgcolor="<% =strHeadCellColor %>"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">
					过去事项：
					</font></b></td>
				</tr>
				<tr>
					<td>
					<% call emitpastEvents %>
					<br>
					<br>
					</td>
				</tr>
			</table>
		</td>
		<td bgcolor="<% =strForumCellColor %>" valign="top" width=100%>
			<table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
				<tr>
					<td valign=top>
						<%if smode = "" then
							emitmonths2 
						else%>
									
						<table bgcolor="<% =strForumCellColor %>" border="0" cellpadding="2" cellspacing="0" width="100%">
							<tr>
								<td align="center" bgcolor="<% =strHeadCellColor %>" nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">
								<% call emitEvents()%>
								</font></b></td>
							</tr>
							<tr>
								<td>
								<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href=events.asp>回到月历首页</a></font>
								</td>
							</tr>
						</table>
						<%end if%>		
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
	'------------------------------------------------------------
	' This function finds the last date of the given month
	'------------------------------------------------------------
	Function GetLastDay(intMonthNum, intYearNum)
		Dim dNextStart
		If CInt(intMonthNum) = 12 Then
			dNextStart = CDate( "1/1/" & intYearNum)
		Else
			dNextStart = CDate(intMonthNum + 1 & "/1/" & intYearNum)
		End If
		GetLastDay = Day(dNextStart - 1)
	End Function
	
	'-------------------------------------------------------------------------
	' This routine prints the individual table divisions for days of the month
	'-------------------------------------------------------------------------
	Sub Write_TD(sValue, sClass)
		Response.Write "		<TD ALIGN='RIGHT' WIDTH=20 HEIGHT=15 VALIGN='BOTTOM' CLASS='" & sClass & "'> " & sValue & "</TD>" & vbCrLf
	End Sub

	Sub Write_TD3(sValue, sClass)
		Response.Write "		<TD HEIGHT='80' WIDTH='14%' ALIGN='left' VALIGN='top' CLASS='" & sClass & "'> <font face=" & strDefaultFontFace & " size=" & strDefaultFontSize  & ">" & sValue & "</font></TD>" & vbCrLf
	End Sub

' This routines presents the form to the user. Depending on the mode, the
' form fields are filled with data or blank.
Sub Show_Form()
	Dim sButtonMsg
	If sMode = "edit" Then
		sButtonMsg = "更新事项"
	Else
		sButtonMsg = "新增事项"
	End If
	
	' Set the title of the page depending on if an error occurred during
	' form submission or not
	If bError = True Then
		sTitle = "送出资料有误"
	Else
		sTitle = "新增或修改事项"
	End If
%>
	 

<table border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td bgcolor="<% =strPopUpBorderColor %>">

	<table border="0" cellspacing="1" cellpadding="4">
	<form NAME="eventfrm" ACTION="<% =sScript %>" METHOD="POST">

		<tr>
			<td bgColor="<% =strPopUpTableColor %>" VALIGN="TOP" ALIGN="right" WIDTH="25%"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">开始日期：</font></td>
			<td bgColor="<% =strPopUpTableColor %>" VALIGN="TOP" ALIGN="LEFT" WIDTH="75%" ><input TYPE="TEXT" SIZE="12" MAXLENGTH="12" NAME="START_DATE" VALUE="<%if Server.HTMLEncode(sStart_Date) <> "" then Response.write Server.HTMLEncode(sStart_Date) else Response.Write(Request.QueryString("date")) end if%>">
			<input type="Checkbox" name="isPrivateEvent" value="1" <% if intPrivateEvent = 1 then response.write "checked" %>>私人事项</td>
		</tr>

		<tr>
			<td bgColor="<% =strPopUpTableColor %>" VALIGN="TOP" ALIGN="RIGHT" WIDTH="30%"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">结束日期：</font></td>
			<td bgColor="<% =strPopUpTableColor %>" VALIGN="TOP" ALIGN="LEFT" WIDTH="70%"><input TYPE="TEXT" SIZE="12" MAXLENGTH="12" NAME="END_DATE" VALUE="<%if Server.HTMLEncode(sEnd_Date) <> "" then Response.write Server.HTMLEncode(sEnd_Date) else Response.Write(Request.QueryString("date")) end if%>"></td>
		</tr>

		<tr>
			<td bgColor="<% =strPopUpTableColor %>" VALIGN="TOP" ALIGN="RIGHT"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">事项名称：</font></td>
			<td bgColor="<% =strPopUpTableColor %>" VALIGN="TOP" ALIGN="LEFT"><input TYPE="TEXT" SIZE="30" MAXLENGTH="100" NAME="EVENT_TITLE" VALUE="<% =Server.HTMLEncode(sEvent_Title) %>"></td>
		</tr>

		<tr>
			<td bgColor="<% =strPopUpTableColor %>" VALIGN="TOP" ALIGN="RIGHT"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">事项内容：</font></td>
			<td bgColor="<% =strPopUpTableColor %>" VALIGN="TOP" ALIGN="LEFT"><textarea COLS="33" ROWS="15" NAME="EVENT_DETAILS" WRAP="PHYSICAL"><% =Server.HTMLEncode(sEvent_Details) %></textarea></td>
		</tr>
		
		<tr>
			<td bgColor="<% =strPopUpTableColor %>" >&nbsp;</td>
			<td bgColor="<% =strPopUpTableColor %>" >
			<input TYPE="HIDDEN" NAME="event" VALUE="<%if sMode = "add" then Response.Write "ADD" else Response.write sMode %>">
			<input TYPE="HIDDEN" NAME="event_id" VALUE="<% =Request.QueryString("EVENT_ID") %>">
			<input TYPE="SUBMIT" VALUE="<% =sButtonMsg %>" id="SUBMIT1" name="SUBMIT1">
			</td>
		</tr>
			</form>
	</table>
	
    </td>
  </tr>
</table>

<% 
End Sub 

Sub Update_Event(sUpdateMode)
	'Dim sErrImage, Rs, sSQL, sTitle, enddate, imgStart_date, imgEnd_date, imgEvent_Title, imgEvent_details
	dim enddate, memberid
	bError = False
	
	' Get the form data into variables
	sStart_Date = Request.Form("START_DATE")
	sEnd_Date = Request.Form("END_DATE")
	sEvent_Title = Request.Form("EVENT_TITLE")
	sEvent_Details = Request.Form("EVENT_DETAILS")
	
	' The error checking routines start here. If any of the fields fails
	' error check, the bError flag is raised, and the corresponding image
	' string is set to the image source
	
	' Check to see if Start date is a valid date
	If NOT IsDate(sStart_Date) Then
		bError = True
		'imgStart_Date = sErrImage
	End If 
	
	' Check if the End date is a valid date provided it is not empty
	If Trim(sEnd_Date) <> "" AND NOT IsDate(sEnd_Date) Then
		bError = True
		'imgEnd_Date = sErrImage
	End If
	
	' Check if the event title field is blank
	If Trim(sEvent_Title) = "" Then
		bError = True
		'imgEvent_Title = sErrImage
	End If 
	
	' Check if the event details field is blank
	If Trim(sEvent_Details) = "" Then
		bError = True
		'imgEvent_Details = sErrImage
	End If 
	
	' Check if the start date and end date correspond to a valid range
	' so that Start date is always prior to the end date
	If IsDate(sStart_Date) AND IsDate(sEnd_Date) Then
		If CDate(sStart_Date) > CDate(sEnd_Date) Then
			bError = True
			'imgStart_Date = sErrImage
			'imgEnd_Date = sErrImage
		End If
	End If

	If bError = False Then
	  		' If the end date is blank, it equals start date
			If Trim(sEnd_Date) <> "" Then
				enddate = CDate(Request.Form("END_DATE"))
			Else
				enddate = CDate(Request.Form("START_DATE"))
			End If		
		
		sSql = "Select Member_ID from " & strMemberTablePrefix & "Members where " & strDBNTSQLName & " = '" & strDBNTUserName & "'"
		dim rs
		set rs = server.CreateObject("adodb.recordset")
		rs.Open sSql, my_Conn
		memberid = rs("Member_ID")
		rs.close
		set rs = nothing
		' Depending on the Mode, open the Recordset in eith Add or Edit mode
	if Request.Form("isPrivateEvent") <> "1" then
		intPrivateEvent = 0
	else
		isPrivateEvent = 1
	end if

		If sUpdateMode = "edit" Then
			sSQL = "UPDATE " & strTablePrefix & "EVENTS SET Start_Date = '" & DateToStr(Request.Form("START_DATE")) & "', End_Date = '" & DateToStr(enddate) & "', Event_Title = '" & chkstring(Request.Form("EVENT_TITLE"), "message") & "', Event_Details = '" & chkstring(Request.Form("EVENT_Details"), "message") & "',PRIVATE = " & intPrivateEvent & ", date_added = '" & DateToStr(now()) & "' WHERE Event_ID=" & Request.Form("Event_ID") 
			my_Conn.Execute sSQL
		Else
			sSQL = "INSERT INTO " & strTablePrefix & "EVENTS (Start_Date, End_Date, Event_Title, Event_Details, Date_Added, Added_by,PRIVATE) Values ('" & DateToStr(Request.Form("START_DATE")) & "', '" & DateToStr(enddate) & "',' " & chkstring(Request.Form("EVENT_TITLE"), "message") & "', '" & chkstring(Request.Form("EVENT_Details"), "message") & "', '" & DateToStr(now()) & "', '" & memberid & "',"  & intPrivateEvent & ")"
			my_Conn.Execute sSQL
		End If
		
		If sUpdateMode = "ADD" Then
			sTitle = "事项已加入<br>感谢你的使用！"
		Else
			sTitle = "事项已更新<br>感谢你的使用！"
		End If
	Else
		If sUpdateMode = "ADD" Then
			sTitle = "<font color=red>发生问题</font><br>"
			sTitle = sTitle & "事项并未加入<br>请重试一遍！"
		Else
			sTitle = "<font color=red>发生问题</font><br>"
			sTitle = sTitle & "事项并未修改<br>请重试一遍！"
		End If
	End If	
End Sub	

function fixQuote(fstring)
	fString = Replace(fString, "'", "''")
	fixquote = fstring
end function

Function emitmonths()

	'start with previous month
	'intthismonth =  intThisMonth - 1

	'show 3 months
	'for i = 0 to 2
		strMonthName = MonthName(intThisMonth)
		datFirstDay = DateSerial(intThisYear, intThisMonth, 1)
		intFirstWeekDay = WeekDay(datFirstDay, vbSunday)
		intLastDay = GetLastDay(intThisMonth, intThisYear)
		
		' Get the previous month and year
		intPrevMonth = intThisMonth - 1
		If intPrevMonth = 0 Then
			intPrevMonth = 12
			intPrevYear = intThisYear - 1
		Else
			intPrevYear = intThisYear	
		End If
		
		' Get the next month and year
		intNextMonth = intThisMonth + 1
		If intNextMonth > 12 Then
			intNextMonth = 1
			intNextYear = intThisYear + 1
		Else
			intNextYear = intThisYear
		End If

		' Get the last day of previous month. Using this, find the sunday of
		' last week of last month
'###################### Added below on 1/11/2001
		if Request.QueryString("month") = "" then
			intLastMonth = DatePart( "m", DateAdd( "m", -1, Date()))
		else
		   	if Request.QueryString("month") = 1 then
		   		intLastMonth = 12
			else
				intLastMonth = Request.QueryString("month") - 1
			end if
		end if	
		if Request.QueryString("year") = "" then
			intPrevYear = DatePart( "yyyy", DateAdd( "m", -1, Date()))
		else
			if Request.QueryString("month") = 1 then
		  		intPrevYear = Request.QueryString("year") - 1
			else
		  		intPrevYear = Request.QueryString("year")
			end if
		end if
'###################### Added above on 1/11/2001

		' Get the last day of previous month. Using this, find the sunday of
		' last week of last month
		LastMonthDate = GetLastDay(intLastMonth, intPrevYear) - intFirstWeekDay + 2
		NextMonthDate = 1

		' Initialize the print day to 1  
		intPrintDay = 1

		
		' These dates are used in the SQL
		dFirstDay = intThisMonth & "/1/" & intThisYear
		dLastDay 	= intThisMonth & "/" & intLastDay & "/" & intThisYear

			sSQL = 	"SELECT event_id, start_date, end_date, event_title, event_details, M_Name, Private FROM FORUM_EVENTS Inner JOIN Forum_Members ON FORUM_EVENTS.added_by = Forum_Members.Member_ID WHERE " & _
						"(Start_Date >='" & DateToStr(dFirstDay) & "' AND Start_Date <= '" & DateToStr(dLastDay) & "') " & _
						"OR " & _
						"(End_Date >='" & DateToStr(dFirstDay) & "' AND End_Date <= '" & DateToStr(dLastDay) & "') " & _
						"OR " & _
						"(Start_Date < '" & DateToStr(dFirstDay) & "' AND End_Date > '" & DateToStr(dLastDay) & "' )"  & _
						"ORDER BY Start_Date"
		'Response.Write sSQL
		'Open the RecordSet with a static cursor. This cursor provides bi-directional navigation
		'Rs.Open sSQL, my_Conn, adOpenStatic, adLockReadOnly, adCmdText
		dim rs
		set rs = server.CreateObject("adodb.recordset")
		
		rs.Open ssql, my_Conn
	%>
					<tr>
						<td valign="top">
	<table ALIGN="CENTER" BORDER="1" CELLSPACING="1" CELLPADDING="4" BGCOLOR="White" BORDERCOLOR="Gray">
	<tr><td>
		<table WIDTH="140" BORDER="0" CELLPADDING="1" CELLSPACING="0" BGCOLOR="#FFFFFF">
			<tr HEIGHT="18" BGCOLOR="Silver">
				<td WIDTH="20" HEIGHT="18" ALIGN="LEFT" VALIGN="MIDDLE"><a HREF="<% =sScript%>?month=<% =IntPrevMonth %>&amp;year=<% =IntPrevYear %>"><img src="<%=strImageURL %>prev.gif" WIDTH="10" HEIGHT="18" BORDER="0" ALT="上一个月"></a></td>
				<td WIDTH="120" COLSPAN="5" ALIGN="CENTER" VALIGN="MIDDLE" CLASS="SOME"><% = strMonthName & " " & intThisYear %></td>
				<td WIDTH="20" HEIGHT="18" ALIGN="RIGHT" VALIGN="MIDDLE"><a HREF="<% =sScript %>?month=<% =IntNextMonth %>&amp;year=<% =IntNextYear %>"><img src="<%=strImageURL %>next.gif" WIDTH="10" HEIGHT="18" BORDER="0" ALT="下一个月"></a></td>
			</tr>
		  <tr>
				<td ALIGN="RIGHT" CLASS="SOME" WIDTH="20" HEIGHT="15" VALIGN="BOTTOM">S</td>
				<td ALIGN="RIGHT" CLASS="SOME" WIDTH="20" HEIGHT="15" VALIGN="BOTTOM">M</td>
				<td ALIGN="RIGHT" CLASS="SOME" WIDTH="20" HEIGHT="15" VALIGN="BOTTOM">T</td>
				<td ALIGN="RIGHT" CLASS="SOME" WIDTH="20" HEIGHT="15" VALIGN="BOTTOM">W</td>
				<td ALIGN="RIGHT" CLASS="SOME" WIDTH="20" HEIGHT="15" VALIGN="BOTTOM">T</td>
				<td ALIGN="RIGHT" CLASS="SOME" WIDTH="20" HEIGHT="15" VALIGN="BOTTOM">F</td>
				<td ALIGN="RIGHT" CLASS="SOME" WIDTH="20" HEIGHT="15" VALIGN="BOTTOM">S</td>
		  </tr>
		  <tr><td HEIGHT="1" ALIGN="MIDDLE" COLSPAN="7"><img src="<%=strImageURL %>line.gif" HEIGHT="1" WIDTH="140" BORDER="0"></td></tr>
		  <%
				' Initialize the end of rows flag to false
				EndRows = False
				Response.Write vbCrLf
				
				' Loop until all the rows are exhausted
			 	Do While EndRows = False
					' Start a table row
					Response.Write "	<TR>" & vbCrLf
					' This is the loop for the days in the week
					For intLoopDay = cSUN To cSAT
						' If the first day is not sunday then print the last days of previous month in grayed font
						If intFirstWeekDay > cSUN Then
							Write_TD LastMonthDate, "NON"
							LastMonthDate = LastMonthDate + 1
							intFirstWeekDay = intFirstWeekDay - 1
						' The month starts on a sunday	
						Else
							' If the dates for the month are exhausted, start printing next month's dates
							' in grayed font
							If intPrintDay > intLastDay Then
								Write_TD NextMonthDate, "NON"
								NextMonthDate = NextMonthDate + 1
								EndRows = True 
							Else
								' If last day of the month, flag the end of the row
								If intPrintDay = intLastDay Then
									EndRows = True
								End If
								
								dToday = CDate(intThisMonth & "/" & intPrintDay & "/" & intThisYear)  
								If NOT Rs.EOF Then
									' Set events flag to false. This means the day has no event in it
									bEvents = False
								  Do While NOT Rs.EOF AND bEvents = False
								  sEName = lcase(RS("M_NAME"))
										' If the date falls within the range of dates in the recordset, then 
										' the day has an event. Make the events flag True
								    If dToday >= strToDate(Rs("Start_Date")) AND dToday <= strToDAte(Rs("End_Date")) AND ((Rs("PRIVATE") <> 1)  OR ( Rs("PRIVATE") = 1 and sEName = lcase(strDBNTUserName) ) and strDBNTUSerName <> "") Then
											' Print the date in a highlighted font
									select case dtoday
									case date()
								      Write_TD "<A HREF=events.asp?date="& Server.URLEncode(dToday) & "&month=" & month(dToday) & "&year=" & year(dToday) & " CLASS='EVENT'> " & intPrintDay & "</A>", "Today"
									case cdate(ddate)
								      Write_TD "<A HREF=events.asp?date="& Server.URLEncode(dToday) & "&month=" & month(dToday) & "&year=" & year(dToday) & " CLASS='EVENT'> " & intPrintDay & "</A>", "Selected"
									case else
								      Write_TD "<A HREF=events.asp?date="& Server.URLEncode(dToday) & "&month=" & month(dToday) & "&year=" & year(dToday) & " CLASS='EVENT'> " & intPrintDay & "</A>", "HL"
									end select
											bEvents = True
										' If the Start date is greater than the date itself, there is no point
										' checking other records. Exit the loop	
								    ElseIf dToday < strToDate(Rs("Start_Date")) Then
											Exit Do
										' Move to the next record
										Else	
									    Rs.MoveNext
										End If
								  Loop
									' Checks for that day
									Rs.MoveFirst
								End If
								
								' If the event flag is not raise for that day, print it in a plain font
								If bEvents = False Then
									select case dtoday
									case date()
									Write_TD "<A HREF=events.asp?date="& Server.URLEncode(dToday) & "&month=" & month(dToday) & "&year=" & year(dToday) & " CLASS='NOEVENT'> " & intPrintDay & "</A>", "TODAY"
									case cdate(ddate)
									Write_TD "<A HREF=events.asp?date="& Server.URLEncode(dToday) & "&month=" & month(dToday) & "&year=" & year(dToday) & " CLASS='NOEVENT'> " & intPrintDay & "</A>", "Selected"
									case else
									Write_TD "<A HREF=events.asp?date="& Server.URLEncode(dToday) & "&month=" & month(dToday) & "&year=" & year(dToday) & " CLASS='NOEVENT'> " & intPrintDay & "</A>", "SOME"
									end select
								End If
							End If 
							
							' Increment the date. Done once in the loop.
							intPrintDay = intPrintDay + 1
						End If
					
					' Move to the next day in the week
					Next
					Response.Write "	</TR>" & vbCrLf
					
				Loop 
				Rs.Close
				set rs = nothing
			%>
		</table>
		</td></tr>
	</table>
						</td>
					</tr>
	<%
		' Get the next month and year
		'intThisMonth = intThisMonth + 1
		'If intThisMonth > 12 Then
		'	intThisMonth = 1
		'	intThisYear = intThisYear + 1
		'Else
		'	intThisYear = intThisYear
		'End If
		'next

end function

function emitEvents()

	if smode = "edit" then
			' Create a recordset and open the given event. If the Record is not found
			' set the mode to ADD
			sSQL = "SELECT start_date, end_date, event_title, event_details, PRIVATE FROM " & strTablePrefix & "EVENTS WHERE Event_ID = " & Request.QueryString("Event_ID")
			set rs = server.CreateObject("adodb.recordset")
			Rs.MaxRecords = 1
			Rs.Open sSQL, my_Conn
	
			If Not Rs.EOF Then
				sStart_Date = strToDate(Rs("Start_Date"))
				sEnd_Date = strToDate(Rs("End_Date"))
				sEvent_Title = Rs("Event_Title")
				sEvent_Details = Rs("Event_Details")
				iEvent_Private = RS("PRIVATE")
			end if
			Response.Write "编辑：" & sEvent_Title
%>
								</td>
							</tr>
							<tr>
								<td <% =strPopUpBorderColor %>>
<%
			if mlev >= 1 then
				Show_Form()
			else
				Response.Write "<font face=" & strDefaultFontFace & " size=" & strDefaultFontSize & " color=red>You need to be a member of " & strForumTitle & " in order to edit an event</font><br>"
				Response.Write "<font face=" & strDefaultFontFace & " size=" & strDefaultFontSize & " color=red>To become a member, <a href=policy.asp>register</a> here.</font>"
			end if
			rs.close
			set rs = nothing
			exit function
	end if

	if smode = "add" then

		Response.Write "新增一件事项" & sEvent_Title
%>
								</td>
							</tr>
							<tr>
								<td>
<%
		if mlev >= 1 then
			Show_Form()
		else
				Response.Write "<font face=" & strDefaultFontFace & " size=" & strDefaultFontSize & " color=red>You need to be a member of " & strForumTitle & " in order to edit an event</font><br>" 
				Response.Write "<font face=" & strDefaultFontFace & " size=" & strDefaultFontSize & " color=red>To become a member, <a href=policy.asp>register</a> here.</font>"
		end if
		exit function
	end if
		'if no date is set, select the events for the whole month, else select events for that day
		If IsEmpty(Request.QueryString("Date")) OR NOT IsDate(Request.QueryString("Date")) Then
			' These dates are used in the SQL

			strMonthName = MonthName(intThisMonth)
			intLastDay = GetLastDay(intThisMonth, intThisYear)
			dFirstDay = intThisMonth & "/1/" & intThisYear
			dLastDay 	= intThisMonth & "/" & intLastDay & "/" & intThisYear

			Response.Write "Events for " & strMonthName
			
			sSQL = 	"SELECT event_id, start_date, end_date, event_title, event_details, PRIVATE, M_Name FROM FORUM_EVENTS Inner JOIN Forum_Members ON FORUM_EVENTS.added_by = Forum_Members.Member_ID WHERE " & _
							"(Start_Date >='" & DateToStr(dFirstDay) & "' AND Start_Date <= '" & DateToStr(dLastDay) & "') " & _
							"OR " & _
							"(End_Date >='" & DateToStr(dFirstDay) & "' AND End_Date <= '" & DateToStr(dLastDay) & "') " & _
							"OR " & _
							"(Start_Date < '" & DateToStr(dFirstDay) & "' AND End_Date > '" & DateToStr(dLastDay) & "' )"  & _
							"ORDER BY Start_Date"
		Else
			
			Response.Write FormatDateTime(dDate, 1) & "的事项"
			
			sSQL = 	"SELECT event_id, start_date, end_date, event_title, event_details, M_Name,PRIVATE FROM FORUM_EVENTS Inner JOIN Forum_Members ON FORUM_EVENTS.added_by = Forum_Members.Member_ID " & _
					"WHERE Start_Date <= '" & DateToStr(dDate) & "' AND End_Date >= '" & DateToStr(dDate) & "' ORDER BY Event_ID "

		End If
%>

								</td>
							</tr>
							<tr>
								<td>
<%		if strCMessage <> "" then Response.Write "<font face=" & strDefaultFontFace & " size=" & strDefaultFontSize & ">" & strCMessage & "</font>"%>
<%		if sTitle <> "" then%>
									<% =sTitle %><br>

								</td>
							</tr>
							<tr>
								<td>
<%		end if

		' Open a record set of schedules
		'Response.Write ssql
		dim rs, strMessage
		Set Rs = Server.CreateObject("ADODB.RecordSet")
		'Response.Write ssql
		Rs.Open sSQL, my_Conn
		
		if rs.eof or rs.bof then
			rs.close
			sSQL = 	"SELECT event_id, start_date, end_date, event_title, event_details,PRIVATE FROM FORUM_EVENTS " & _
					"WHERE Start_Date <= '" & DateToStr(dDate) & "' AND End_Date >= '" & DateToStr(dDate) & "' ORDER BY Event_ID "
			Rs.Open sSQL, my_Conn
			MemberErased = true
		end if
		If NOT Rs.EOF Then

			Do While NOT Rs.EOF
			sEName = lcase(rs("M_NAME"))
		if (rs("PRIVATE") <> 1) or (rs("PRIVATE") = 1 and sEName = lcase(strDBNTUSerName)) and strDBNTUSerName <> "" then%>	
						<br>
						<table align="center" width="95%" cellpadding="0" cellspacing="0" border="0" bgcolor="gray">
							<tr>
								<td>
									<table WIDTH="100%" CELLSPACING="1" BORDER="0" CELLPADDING="3" align="center">
										<tr>
											<td align="left" bgcolor="<% =strHeadCellColor %>"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">
											<%Response.Write Trim(Rs("Event_Title"))
											strMessage = Rs("Event_Details")%>
											<%
											if not(memberErased) then
												Response.Write " - 加入者：" & RS("M_Name")
											else
												Response.Write " - 由已离开会员加入"
											end if
											%>											</font></b></td>
											<%If ((mlev >= 1) and (lcase(RS("M_Name")) = lcase(strDBNTUserName))) or mlev >=3 Then%>
											<%'If mlev > 2 or strDBNTUserName = RS("M_Name") Then%>
											<td align="center" BGCOLOR="<%= strCategoryCellColor%>" width="50">
												<a HREF="events.asp?mode=edit&Event_ID=<%= Rs("Event_ID")%>"><img src="<%=strImageURL %>icon_pencil.gif" BORDER="0" WIDTH="12" HEIGHT="12"></a>
												<a HREF="events.asp?mode=delete&Event_ID=<%=Rs("Event_ID")%>&date=<%= server.HTMLEncode(request.QueryString("date"))%>"><img src="<%=strImageURL %>icon_trashcan.gif" BORDER="0" WIDTH="12" HEIGHT="12"></a>
											</td>
											<%End If%>
										</tr>
										<tr>
											<td height="30" bgcolor="white" <%If mlev >=3 or lcase(strDBNTUserName) = lcase(RS("M_Name")) Then Response.Write "colspan=2"%>>
												<%
												' If the event lasts more than one day, indidate the start and end dates
												If Rs("Start_Date") <> Rs("End_Date") Then
													Response.Write "<FONT FACE='宋体, Arial' SIZE=1 COLOR='Gray'>开始於：" & strToDate(Rs("Start_Date")) & vbCrLf
													Response.Write "<BR>终止於：" & strToDate(Rs("End_Date")) & vbCrLf
													Response.Write "</FONT><P>"
												else
													Response.Write "<FONT FACE='宋体, Arial' SIZE=1 COLOR='Gray'>开始於：" & strToDate(Rs("Start_Date")) & vbCrLf
													Response.Write "</FONT><P>"												
												End If
												'Response.Write Replace(Rs("Event_Details") & " ", vbCrLf, "<BR>")
												%>
												<font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% =formatStr(rs("event_details")) %><a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" height="15" width="15" border="0" align="right" alt="回到页首"></a></font></td>
											</td>
										</tr>
									</table>
								</td>
							</tr>
						</table>
			
			<br>
			<% end if
				Rs.MoveNext
			Loop
		else
			if Request.QueryString("date") = "" then
				Response.Write "<br><font face=" & strDefaultFontFace & " size=" & strDefaultFontSize & ">这个月没有任何事项。</font>"
			else
				Response.Write "<br><font face=" & strDefaultFontFace & " size=" & strDefaultFontSize & ">今天没有任何事项。</font>"
			end if
		End If

	rs.close
	set rs = nothing
end function

function emitupcomingevents
	dim rs
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	strSql = "SELECT start_date, event_title, M_Name,PRIVATE FROM FORUM_EVENTS Inner JOIN Forum_Members ON FORUM_EVENTS.added_by = Forum_Members.Member_ID WHERE start_date >= '" & DateToStr(date()) & "' and start_date < '" & DateToStr(DateAdd("d",30,date())) & "' Order by start_date, event_id ASC"
		
	rs.Open strSql, my_Conn
	do until rs.EOF
	sEName = lcase(rs("M_NAME"))
		if (rs("PRIVATE") <> 1) or (rs("PRIVATE") = 1 and sEName = lcase(strDBNTUSerName)) and strDBNTUSerName <> "" then
			Response.Write "<font face=" & strDefaultFontFace & " size=" & strDefaultFontSize & "><a href=events.asp?date=" & strToDate(rs("Start_Date")) & ">" & rs("Event_Title") & "</a></font><br>"
		end if
		rs.MoveNext
	loop
	rs.Close
	set rs = nothing
end function

function emitpastEvents
	dim rs
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	strSql = "SELECT start_date, event_title, M_Name,PRIVATE FROM FORUM_EVENTS Inner JOIN Forum_Members ON FORUM_EVENTS.added_by = Forum_Members.Member_ID WHERE start_date < '" & DateToStr(date()) & "' and start_date > '" & DateToStr(DateAdd("d",-30,date())) & "' Order by start_date desc"
		
	rs.Open strSql, my_Conn
	do until rs.EOF
	sEName = lcase(rs("M_NAME"))

		if (rs("PRIVATE") <> 1) or (rs("PRIVATE") = 1 and sEName = lcase(strDBNTUSerName)) and strDBNTUSerName <> "" then
			Response.Write "<font face=" & strDefaultFontFace & " size=" & strDefaultFontSize & "><a href=events.asp?date=" & strToDate(rs("Start_Date")) & ">" & rs("Event_Title") & "</a></font><br>"
		end if
		rs.MoveNext
	loop
	rs.Close
	set rs = nothing
end function

Function emitmonths2()

	'start with previous month
	'intthismonth =  intThisMonth - 1

	'show 3 months
	'for i = 0 to 2
		strMonthName = MonthName(intThisMonth)
		datFirstDay = DateSerial(intThisYear, intThisMonth, 1)
		intFirstWeekDay = WeekDay(datFirstDay, vbSunday)
		intLastDay = GetLastDay(intThisMonth, intThisYear)
		
		' Get the previous month and year
		intPrevMonth = intThisMonth - 1
		If intPrevMonth = 0 Then
			intPrevMonth = 12
			intPrevYear = intThisYear - 1
		Else
			intPrevYear = intThisYear	
		End If
		
		' Get the next month and year
		intNextMonth = intThisMonth + 1
		If intNextMonth > 12 Then
			intNextMonth = 1
			intNextYear = intThisYear + 1
		Else
			intNextYear = intThisYear
		End If

		' Get the last day of previous month. Using this, find the sunday of
		' last week of last month
'###################### Added below on 1/11/2001
		if Request.QueryString("month") = "" then
			intLastMonth = DatePart( "m", DateAdd( "m", -1, Date()))
		else
		   	if Request.QueryString("month") = 1 then
		   		intLastMonth = 12
			else
				intLastMonth = Request.QueryString("month") - 1
			end if
		end if	
		if Request.QueryString("year") = "" then
			intPrevYear = DatePart( "yyyy", DateAdd( "m", -1, Date()))
		else
			if Request.QueryString("month") = 1 then
		  		intPrevYear = Request.QueryString("year") - 1
			else
		  		intPrevYear = Request.QueryString("year")
			end if
		end if
'###################### Added above on 1/11/2001

		' Get the last day of previous month. Using this, find the sunday of
		' last week of last month
		LastMonthDate = GetLastDay(intLastMonth, intPrevYear) - intFirstWeekDay + 2
		NextMonthDate = 1

		' Initialize the print day to 1  
		intPrintDay = 1

		
		' These dates are used in the SQL
		dFirstDay = intThisMonth & "/1/" & intThisYear
		dLastDay 	= intThisMonth & "/" & intLastDay & "/" & intThisYear

			sSQL = 	"SELECT event_id, start_date, end_date, event_title, event_details, M_Name,PRIVATE FROM FORUM_EVENTS Inner JOIN Forum_Members ON FORUM_EVENTS.added_by = Forum_Members.Member_ID WHERE " & _
						"(Start_Date >='" & DateTostr(dFirstDay) & "' AND Start_Date <= '" & DateTostr(dLastDay) & "') " & _
						"OR " & _
						"(End_Date >='" & DateTostr(dFirstDay) & "' AND End_Date <= '" & DateTostr(dLastDay) & "') " & _
						"OR " & _
						"(Start_Date < '" & DateTostr(dFirstDay) & "' AND End_Date > '" & DateTostr(dLastDay) & "' )"  & _
						"ORDER BY Start_Date"
		'Response.Write sSQL
		'Open the RecordSet with a static cursor. This cursor provides bi-directional navigation
		'Rs.Open sSQL, my_Conn, adOpenStatic, adLockReadOnly, adCmdText
		dim rs
		Set Rs = Server.CreateObject("ADODB.RecordSet")
		rs.Open ssql, my_Conn
	%>
	<table WIDTH="99%" ALIGN="CENTER" BORDER="1" CELLSPACING="1" CELLPADDING="4" BGCOLOR="White" BORDERCOLOR="Gray">
			<tr>
				<td WIDTH="10%" HEIGHT="10" ALIGN="center" VALIGN="MIDDLE" bgcolor="<% =strForumCellColor %>"><a HREF="<% =sScript%>?month=<% =IntPrevMonth %>&amp;year=<% =IntPrevYear %>"><img src="<%=strImageURL %>prev.gif" WIDTH="10" HEIGHT="18" BORDER="0" ALT="上一个月"></a></td>
				<td width="80%" colspan="5" align="center" valign="middle" bgcolor="<% =strCategoryCellColor %>"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">
					<% = strMonthName & " " & intThisYear %>
				</td>
				<td WIDTH="10%" HEIGHT="10" ALIGN="center" VALIGN="MIDDLE" bgcolor="<% =strForumCellColor %>"><a HREF="<% =sScript %>?month=<% =IntNextMonth %>&amp;year=<% =IntNextYear %>"><img src="<%=strImageURL %>next.gif" WIDTH="10" HEIGHT="18" BORDER="0" ALT="下一个月"></a></td>
			</tr>
		  <tr>
				<td HEIGHT="20" WIDTH="14%" ALIGN="center" VALIGN="middle"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">星期日</td>
				<td HEIGHT="20" WIDTH="14%" ALIGN="center" VALIGN="middle"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">星期一</td>
				<td HEIGHT="20" WIDTH="14%" ALIGN="center" VALIGN="middle"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">星期二</td>
				<td HEIGHT="20" WIDTH="14%" ALIGN="center" VALIGN="middle"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">星期三</td>
				<td HEIGHT="20" WIDTH="14%" ALIGN="center" VALIGN="middle"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">星期四</td>
				<td HEIGHT="20" WIDTH="14%" ALIGN="center" VALIGN="middle"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">星期五</td>
				<td HEIGHT="20" WIDTH="14%" ALIGN="center" VALIGN="middle"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">星期六</td>
		  </tr>
		  <%
				' Initialize the end of rows flag to false
				EndRows = False
				Response.Write vbCrLf
				
				' Loop until all the rows are exhausted
			 	Do While EndRows = False
					' Start a table row
					Response.Write "	<TR>" & vbCrLf
					' This is the loop for the days in the week
					For intLoopDay = cSUN To cSAT
						' If the first day is not sunday then print the last days of previous month in grayed font
						If intFirstWeekDay > cSUN Then
							Write_TD3 LastMonthDate, "NON2"
							LastMonthDate = LastMonthDate + 1
							intFirstWeekDay = intFirstWeekDay - 1
						' The month starts on a sunday	
						Else
							' If the dates for the month are exhausted, start printing next month's dates
							' in grayed font
							If intPrintDay > intLastDay Then
								Write_TD3 NextMonthDate, "NON2"
								NextMonthDate = NextMonthDate + 1
								EndRows = True 
							Else
								' If last day of the month, flag the end of the row
								If intPrintDay = intLastDay Then
									EndRows = True
								End If
								
								dToday = CDate(intThisMonth & "/" & intPrintDay & "/" & intThisYear)  
								If NOT Rs.EOF Then
									' Set events flag to false. This means the day has no event in it
									bEvents = False
								  Do While NOT Rs.EOF AND bEvents = False
										' If the date falls within the range of dates in the recordset, then 
										' the day has an event. Make the events flag True
								    If dToday >= strToDate(Rs("Start_Date")) AND dToday <= strToDate(Rs("End_Date"))  then 
									  if dtoday = date() then
											' Print the date in a highlighted font
										Write_TD2 "<A HREF=events.asp?date="& Server.URLEncode(dToday) & "&month=" & month(dToday) & "&year=" & year(dToday) & " CLASS='EVENT'> " & intPrintDay & "</A>", "TODAY", dToday
									  else
											' Print the date in a highlighted font
										Write_TD2 "<A HREF=events.asp?date="& Server.URLEncode(dToday) & "&month=" & month(dToday) & "&year=" & year(dToday) & " CLASS='EVENT'> " & intPrintDay & "</A>", "HL", dToday
								      end if
											bEvents = True
										' If the Start date is greater than the date itself, there is no point
										' checking other records. Exit the loop	
								    ElseIf dToday < strToDate(Rs("Start_Date")) Then
											Exit Do
										' Move to the next record
										Else	
									    Rs.MoveNext
										End If
								  Loop
									' Checks for that day
									Rs.MoveFirst
								End If
								
								' If the event flag is not raise for that day, print it in a plain font
								If bEvents = False Then
									if dtoday = date() then
									Write_TD3 "<A HREF=events.asp?date="& Server.URLEncode(dToday) & "&month=" & month(dToday) & "&year=" & year(dToday) & " CLASS='NOEVENT'> " & intPrintDay & "</A>", "TODAY"
									else
									Write_TD3 "<A HREF=events.asp?date="& Server.URLEncode(dToday) & "&month=" & month(dToday) & "&year=" & year(dToday) & " CLASS='NOEVENT'> " & intPrintDay & "</A>", "SOME"
									end if
								End If
							End If 
							
							' Increment the date. Done once in the loop.
							intPrintDay = intPrintDay + 1
						End If
					
					' Move to the next day in the week
					Next
					Response.Write "	</TR>" & vbCrLf
					
				Loop 
				Rs.Close
				set rs = nothing
			%>
		</table>
		</td>
	</tr>
	<tr>
		<td>
	<%
		' Get the next month and year
		'intThisMonth = intThisMonth + 1
		'If intThisMonth > 12 Then
		'	intThisMonth = 1
		'	intThisYear = intThisYear + 1
		'Else
		'	intThisYear = intThisYear
		'End If
		'next

end function

	'-------------------------------------------------------------------------
	' This routine prints the individual table divisions for days of the month
	'-------------------------------------------------------------------------
	Sub Write_TD2(sValue, sClass, dDate)
		dim rsEvents, sETitle
		set rsevents = server.CreateObject("adodb.recordset")
		Response.Write "		<TD HEIGHT='80' WIDTH='14%' ALIGN='left' VALIGN='top' CLASS='" & sClass & "'> " & sValue 
			sSQL = 	"SELECT event_id, start_date, end_date, event_title, event_details, M_Name,PRIVATE FROM FORUM_EVENTS Inner JOIN Forum_Members ON FORUM_EVENTS.added_by = Forum_Members.Member_ID " & _
				"WHERE Start_Date <= '" & DateTostr(dDate) & "' AND End_date >= '" & DateTostr(dDate) & "' ORDER BY Event_ID "
		
		'response.write sSQL
		rsEvents.Open sSQL, my_Conn
		do while not(rsevents.EOF)
			sETitle = rsEvents("event_title")
			sEName = lcase(rsEvents("M_NAME"))
			if len(sETitle) > 14 then
				sETitle = mid(sETitle,1,15) 
			end if
			if (rsEvents("PRIVATE") <> 1) or (rsEvents("PRIVATE") = 1 and sEName = lcase(strDBNTUSerName)) and strDBNTUSerName <> "" then
				Response.Write "<br><A HREF=events.asp?date="& Server.URLEncode(dToday) & ">" & sETitle & "</a>"
			end if
			rsEvents.MoveNext
		loop
		Response.Write "</TD>" & vbCrLf
		rsEvents.Close
		set rsevents = nothing
	End Sub

%>
