
<%
'strConnString = "driver={SQL Server};server=flyagaric;uid=sa;pwd=wynnston;database=HRForum"
strTablePrefix = "FORUM_"
function PollMentor_GetDatabaseConn()
	Dim oRet
	Dim strDSN
	Set oRet = Server.CreateObject	("ADODB.Connection")
	oRet.Open strConnString	
	Set PollMentor_GetDatabaseConn = oRet
End function

function PollMentor_GetTitle()
	PollMentor_GetTitle = "民意调查"	
End function


'''TODO for you! Configuration:
''''3. Some ads if you'd like

function FAQ_GetAd(nNumber)
	Select Case nNumber
		Case 1
			FAQ_GetAd = ""
		Case 2
			FAQ_GetAd = ""
		Case 3
			FAQ_GetAd = ""
	End Select	
End function

function PollMentor_TryToVote( sID, nNumber )
	Dim sRet, strSQL
	Dim oConn
	Dim strTrue
	Dim strFalse
	sRet = "Huw"
	If strDBType = "sqlserver" Then
		strTrue = 	"1"
		strFalse = "0"
	Else
		strTrue = "True"
		strFalse = "False"
	End If
	
	Set oConn = PollMentor_GetDatabaseConn()
	'Get real id...
	Dim oRS
	If sID = -1 Then
		Set oRS = oConn.Execute("select id from " & strTablePrefix & "poll where active=" & strTrue)
		sID = oRS("id").Value
		oRS.Close
		Set oRS= Nothing
	End If
	
	If PollMentor_CanUserVote( oConn, sID ) = False Then
		sRet = "一个人只能投一票！<br>* 感谢你神圣的一票 *<br>"
	Else
		strSQL = "update " & strTablePrefix & "poll set count" & nNumber & " = count" & nNumber & " +1 where "
		If nNumber =-1 Then
			strSQL = strSQL & " active=" & strTrue
		Else
			strSQL = strSQL & " id=" & sID
		End If

		oConn.Execute strSQL
			If strDBType = "sqlserver" Then
				sTime = " getdate() "
			Else
				sTime = "#" & Now() & "#"
			End If

		oConn.Execute "insert into " & strTablePrefix & "votelog(poll_id, ip,datum) values(" & sID & ",'" & remoteIP() & "'," & sTime & ")"
		sRet = "Thanks for voting"
	End If
	oConn.Close
	Set oConn = Nothing
	PollMentor_TryToVote=sRet
End function


function PollMentor_CanUserVote( oConn, sID ) '
	'Check of user already has voted within 24 hours?
	'If so then no voting can be done...
	' Here's your chance to display some other content
	'1. Check IP address
	Dim strSQL, sTime, oRS
	
	If strDBType = "sqlserver" Then
		sTime = " dateadd(day,-1,getdate()) "
	Else
			sTime = "#" & DateAdd( "d", -1, Now() ) & "#"
	End If
	strSQL = "select id from " & strTablePrefix & "votelog where poll_id=" & sID & " AND datum > " & sTime & " AND ip='" & remoteIP() & "'"
	'Response.Write strSQL
	Set oRS = oConn.Execute(strSQL)	
	If oRS.EOF Then
		PollMentor_CanUserVote = True
	Else
		PollMentor_CanUserVote = False
	End If
	oRS.Close
	Set oRS = Nothing
End function

function PollMentor_GetPollInfo ( ByVal nID, ByRef sTitle, ByRef sQuestion, ByRef vAnswers, ByRef vCount )
	Dim sRet, strSQL
	Dim oConn, oRS, nCount
	Dim strTrue
	Dim strFalse

		if strDBType = "access" then
				strTrue = 	True
				strFalse = false
		else
		strTrue = 	"1"
		strFalse = "0"
		end if
	Set oConn = PollMentor_GetDatabaseConn()
	
	strSQL = "select * from " & strTablePrefix & "poll where " 
	If nID = -1 Then
		strSQL = strSQL & " active=" & strTrue
	Else
		strSQL = strSQL & " id=" & nID
	End If
	Set oRS = oConn.Execute(strSQL)
	If oRS.EOF Then
		PollMentor_GetPollInfo = False
	Else
		sTitle=PollMentor_GetTitle()
		sQuestion=oRS("question").Value
		For nCount=1 To 8
			vAnswers(nCount)=oRS("answer" & CStr(nCount)).Value
			vCount(nCount)=oRS("count" & CStr(nCount)).Value
		Next
		PollMentor_GetPollInfo  = True
	End If
	oRS.Close
	Set oRS = Nothing
	oConn.Close
	Set oConn = Nothing
End function
function remoteIP()	
	
	remoteIP = Request.ServerVariables("HTTP_X_FORWARDED_FOR")	
	if remoteIP = "" then		
		remoteIP = Request.ServerVariables("REMOTE_ADDR")	
	end if
end function


%>