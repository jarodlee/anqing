<!--#INCLUDE FILE="config.asp" -->
<!--#INCLUDE FILE="inc_functions.asp" -->
<!--#INCLUDE FILE="inc_top.asp" -->
<% 
set rs = Server.CreateObject("ADODB.RecordSet")

err_Msg = ""
ok = "" 


if Request.Form("Method_Type") = "Topic" then

	'## Forum_SQL
	strSql = "SELECT MEMBER_ID, M_LEVEL, M_EMAIL, M_NAME"
	if strAuthType = "nt" then
		strSql = strSql & ", M_USERNAME "
	end if
	strSql = strSql & ", M_PASSWORD "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	if strAuthType = "nt" then
		strSql = strSql & " WHERE M_USERNAME = '" & session("userid") & "'"
	else
		if strAuthType = "db" then
			strSql = strSql & " WHERE M_NAME = '" & Request.Form("UserName") & "' "
			strSql = strSql & " AND   M_PASSWORD = '" & Request.Form("Password") &"'"
		end if
	end if

	set rs = my_Conn.Execute (strSql)
	
	if rs.BOF or rs.EOF then '##  Invalid Password
		Go_Result "Invalid UserName or Password", 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
		Response.End
	Else
	
	
strSql = "SELECT MEMBER_ID, M_NAME, M_PMRECEIVE, M_PMEMAIL"
strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
strSql = strSql & " WHERE M_NAME = '" & Request.Form("sendto") & "' "
set rsName = my_Conn.Execute (strSql)

if rsName.BOF or rs.EOF then '##  no one registered
Go_Result "抱歉！查无此人！", 0

%>
<!--#INCLUDE FILE="inc_footer.asp" -->

<% else

if rsName("M_PMRECEIVE") = "0" then 
Go_Result "抱歉！该名会员不希望收到悄悄话讯息...", 0

%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<% end if
end if

		txtMessage = ChkString(Request.Form("Message"),"message")
		txtSubject = ChkString(Request.Form("Subject"),"title")
		FILEZ = Request.Form("FILEZ")

		if txtMessage = " " then
			Go_Result "你必须填写内容", 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
		end if
		if Request.Form("sendto") = "" then
			Go_Result "你必须填写收件人！", 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
		end if
		if txtSubject = " " then
			Go_Result "你必须填写标题!", 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
		end if         
		if Request.Form("sig") = "yes" and GetSig(Request.Form("UserName")) <> "" then
		     txtMessage = txtMessage & vbCrLf & vbCrLf & ChkString(GetSig(Request.Form("UserName")), "signature" )
		end if
		
				
		if Request.Form("rmail") <> "1" then
			TF = "0"
		Else 
			TF = "1"
		end if
		
		'## Forum_SQL - Add new private message
		
		strSql = "INSERT INTO " & strTablePrefix & "PM ("
		strSql = strSql & " M_SUBJECT"
		strSql = strSql & ", M_MESSAGE"
		strSql = strSql & ", M_TO"
		strSql = strSql & ", M_FROM"
		strSql = strSql & ", M_SENT"
		strSql = strSql & ", M_MAIL"
   		strSql = strSql & ", M_READ"
		strSql = strSql & ", M_OUTBOX"
		strSql = strSql & ") VALUES ("
		strSql = strSql & " '" & txtSubject & "'"
		strSql = strSql & ", '" & txtMessage & "'"
		strSql = strSql & ", " & rsName("MEMBER_ID")
		strSql = strSql & ", " & rs("MEMBER_ID")
		strSql = strSql & ", '" & DateToStr(strForumTimeAdjust) & "'"
		strSql = strSql & ", " & TF
		strSql = strSql & ", " & "0"
		if request.cookies("paging")("outbox") = "double" or request.cookies("paging")("outbox") = "single" then
			strSql = strSql & ", " & 1 & ")"
		else
			strSql = strSql & ", " & 0 & ")"
		end if

		my_Conn.Execute (strSql)

		if strEmail = "1" then
			if rsName("M_PMEMAIL") = "1" then
				DoReplyEmail Request.Form("sendto")
                	end if
		end if

		if Err.description <> "" then 
			err_Msg = "发生一个错误 →  " & Err.description
		Else
			err_Msg = "更新完成"
		end if


		Go_Result err_Msg, 1
%>
<!--#INCLUDE FILE="inc_footer.asp" -->

<% 
		Response.End
	end if	
end if

if Request.Form("Method_Type") = "Reply" or Request.Form("Method_Type") = "ReplyQuote" or Request.Form("Method_Type") = "Forward" then

if Request.Form("Method_Type") = "Forward" then
'## Forum_SQL Convert MemberName
strSql = "SELECT MEMBER_ID, M_Name,M_PMEMAIL"
strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
strSql = strSql & " WHERE M_NAME = '" & Request.Form("sendto") & "' "

set rsName = my_Conn.Execute (strSql)

end if


	'## Forum_SQL
	strSql = "SELECT MEMBER_ID, M_LEVEL, M_EMAIL, M_NAME, M_PMEMAIL"
	if strAuthType = "nt" then
		strSql = strSql & ", M_USERNAME "
	end if
	strSql = strSql & ", M_PASSWORD "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	if strAuthType = "nt" then
		strSql = strSql & " WHERE M_USERNAME = '" & session("userid") & "'"
	else
		if strAuthType = "db" then
			strSql = strSql & " WHERE M_NAME = '" & Request.Form("UserName") & "' "
			strSql = strSql & " AND   M_PASSWORD = '" & Request.Form("Password") &"'"
		end if
	end if

	set rsReply = my_Conn.Execute (strSql)




	if rsReply.BOF or rsReply.EOF then '##  Invalid Password
		Go_Result "Invalid UserName or Password", 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
		Response.End
	Else
	
	
	    if Request.Form("Method_Type") = "Forward" then
		txtRE = "FWD: "
	    else 	
	    	txtRE = "RE: "	
	    end if
		txtMessage = ChkString(Request.Form("Message"),"message")
		txtSubject = ChkString(Request.Form("R_SUBJECT"),"title")
		FILEZ = Request.Form("FILEZ")

		if txtMessage = " " then
			Go_Result "你必须填写内容", 0
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
			Response.End
		end if
		
		if Request.Form("sig") = "yes" and GetSig(Request.Form("UserName")) <> "" then
		     txtMessage = txtMessage & vbCrLf & vbCrLf & ChkString(GetSig(Request.Form("UserName")), "signature" )
		end if
		if Request.Form("rmail") <> "1" then
			TF = "0"
		Else 
			TF = "1"
		end if
		
		'## Forum_SQL - Add new private message
		
		strSql = "INSERT INTO " & strTablePrefix & "PM ("
		strSql = strSql & " M_SUBJECT"
		strSql = strSql & ", M_MESSAGE"
		strSql = strSql & ", M_TO"
		strSql = strSql & ", M_FROM"
		strSql = strSql & ", M_SENT"
		strSql = strSql & ", M_MAIL"
   		strSql = strSql & ", M_READ"
		strSql = strSql & ", M_OUTBOX"
		strSql = strSql & ") VALUES ("
		strSql = strSql & " '" & txtRE + txtSubject & "'"
		strSql = strSql & ", '" & txtMessage & "'"
	If Request.Form("Method_Type") = "Forward" then
		strSql = strSql & ", " & rsName("MEMBER_ID")
	else	
		strSql = strSql & ", " & Request.Form("REPLY_ID")
	end if			
		strSql = strSql & ", " & rsReply("MEMBER_ID")
		strSql = strSql & ", '" & DateToStr(strForumTimeAdjust) & "'"
		strSql = strSql & ", " & TF
		strSql = strSql & ", " & "0"
		if request.cookies("paging")("outbox") = "double" or request.cookies("paging")("outbox") = "single" then 
			strSql = strSql & ", " & 1 & ")"
		else
			strSql = strSql & ", " & 0 & ")"
		end if

		my_Conn.Execute (strSql)

		if strEmail = "1" then
			If Request.Form("Method_Type") = "Forward" then
				if rsName("M_PMEMAIL") = "1" then
					DoReplyEmail Request.Form("sendto")
                		end if
			else
				if rsReply("M_PMEMAIL") = "1" then
				DoReplyEmail Request.Form("sendto")
                	end if
		end if
		end if

		if Err.description <> "" then 
			err_Msg = "发生一个错误 →  " & Err.description
		Else
			err_Msg = "更新完成"
		end if


		Go_Result err_Msg, 1
%>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
		Response.End
	end if	
end if

sub Go_Result(str_err_Msg, boolOk)
%>
<table border="0" width="100%">
  <tr>
	<td width="33%" align="left"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
	<img src="<%=strImageURL %>icon_folder_open.gif" border="0">&nbsp;<a href="default.asp">返回论坛首页</a><br>
<% if Request.form("Method_Type") = "Topic" or _
	Request.form("Method_Type") = "ReplyQuote" or _
	Request.form("Method_Type") = "Reply" then %>
	
<% end if %>
<% if Request.form("Method_Type") = "Reply" then %>
	<img src="<%=strImageURL %>icon_blank.gif" border="0"><img src="<%=strImageURL %>icon_bar.gif" border="0"><img src="<%=strImageURL %>icon_folder_open_topic.gif" border="0">&nbsp;<a href="<% =Request.Form("refer") %>"><% =Request.Form("Topic_Title")%></a>
<% end if %>
    </font></td>
  </tr>
</table>

<%	if boolOk = 1 then %>
<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">
<%

%>
</font></p>
<meta http-equiv="Refresh" content="0; URL=pm_view.asp">

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">
<%
		select case Request.Form("Method_Type")

			case "Topic"
				Response.Write("你的讯息已送出")
			case "Reply"
				Response.Write("你的回复已送出")
			case "ReplyQuote"
				Response.Write("你的回复已送出")
			case "Forward"
				Response.Write("你的讯息已转送")
			case "Edit"
				Response.Write("感谢你的参与！")
			case else
				Response.Write("祝你有愉快的一天！")

		end select
%>
</font></p>

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="pm_view.asp">返回悄悄话收件箱</font></a></p>

<!--#INCLUDE FILE="inc_footer.asp" -->
<%		Response.End %>
<%	Else %>
<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">发生问题！</font></p>

<p align="center"><font color="red" size="<% =strHeaderFontSize %>"><% =str_err_Msg %></font></p>

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript:history.go(-1)">返回修正问题。</a></font></p>

<!--#INCLUDE FILE="inc_footer.asp" -->
<%		Response.End %>
<%
	end if 
End Sub

sub DoReplyEmail(PostedBy)
	'## Emails all users who wish to receive notification of private message
	

	'## Forum_SQL
	strSql = " SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.M_EMAIL, " & strMemberTablePrefix & "MEMBERS.M_PMEMAIL, " & strMemberTablePrefix & "PM.M_SUBJECT "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " INNER JOIN " & strTablePrefix & "PM "
	strSql = strSql & " ON         " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "PM.M_TO "
	strSql = strSql & " ORDER BY " & strTablePrefix & "PM.M_SENT  DESC" 

	set rsTopicAuthor = my_Conn.Execute (strSql)
	
	
		strRecipientsName = rsTopicAuthor("M_NAME")
		strRecipients = rsTopicAuthor("M_EMAIL")
		strSubject = strForumTitle & " - 新的悄悄话讯息"
		strMessage = "Hello " & rsTopicAuthor("M_NAME") & vbCrLf & vbCrLf
	strMessage = strMessage & Request.Cookies(strUniqueID & "User")("Name") & "  从 " & strForumTitle & " 发了一条悄悄话讯息给你。" & vbCrLf
		if Request.Form("Subject") = "" then
		strMessage = strMessage & "是关於 - " & rsTopicAuthor("M_SUBJECT") & vbCrLf & vbCrLf
		else
		strMessage = strMessage & "标题是 - " & Request.Form("Subject") & vbCrLf & vbCrLf
		end if
		strMessage = strMessage & "你可以到这观看讯息 " & Left(Request.Form("refer"), InstrRev(Request.Form("refer"), "/")) & vbCrLf
%>
	<!--#INCLUDE FILE="inc_mail.asp" -->
<% end sub %>
