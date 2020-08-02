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
<table width="100%" border=0>
  <tr>
    <td>
    <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><img src="<%=strImageURL %>icon_folder_open.gif" border="0">&nbsp;<a href="default.asp">返回论坛首页</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" border="0">&nbsp;<a href="policy.asp">服务条款和声明</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" border="0">&nbsp;<% =strForumTitle %>会员注册</font>
    </td>
  </tr>
</table>
<%
if Request.QueryString("mode") <> "DoIt" then
	if strAuthType = "nt" and ChkAccountReg = "1" then
%>
<p align="center"><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>><b>这个帐号不　要注册</b></font></p>
<table align="center">
  <tr>
    <td>
    <ul>
      <li>此NT使用者帐号已经注册</li>
    </ul>
    </td>
  </tr>
</table>
<!--#INCLUDE FILE="inc_footer.asp" -->
<%
		Response.End 
	end if			

	call ShowForm

else
	Err_Msg = ""

	If strAutoLogon <> 1 then	
		if Request.Form("Name") = "" then 
			Err_Msg = Err_Msg & "<li>你必须输入姓名</li>"
		end if
	end if

	'## Forum_SQL
	strSql = "SELECT M_NAME FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE M_NAME = '" & Trim(Request.Form("Name")) &"'"

	set rs = my_Conn.Execute (strSql)

	if rs.BOF and rs.EOF then 
		'## Do Nothing
	else 
		Err_Msg = Err_Msg & "<li>这个用户名已有人使用，请选择其他名称</li>"
	end if
	rs.close
	set rs = nothing

	if (Instr(Request.Form("Name"), ">") > 0 ) or (Instr(Request.Form("Name"), "<") > 0) then
		Err_Msg = Err_Msg & "<li> &gt; 和 &lt; 不能出现在用户名里，请重新输入</li>"
	end if

	'## NT authentication no additional password needed
	if strAuthType = "db" then
		if Request.Form("Password") = "" then 
			Err_Msg = Err_Msg &  "<li>你必须输入密码</li>"
		end if
		if Len(Request.Form("Password")) > 25 then 
			Err_Msg = Err_Msg & "<li>密码不能超过25个字符</li>" 
		end if
		if Request.Form("Password") <> Request.Form("Password2") then 
			Err_Msg = Err_Msg & "<li>两次输入的密码不同</li>"
		end if
	end if

	If strAutoLogon <> 1 then
		if Request.Form("Email") = "" then 
			Err_Msg = Err_Msg & "<li>必须填写邮件地址</li>"
		end if
		if EmailField(Request.Form("Email")) = 0 then 
			Err_Msg = Err_Msg & "<li>必须填写正确的邮件地址</li>"
		end if
	end if

	if (lcase(left(Request.Form("Homepage"), 7)) <> "http://") and (lcase(left(Request.Form("Homepage"), 8)) <> "https://") and (Request.Form("Homepage") <> "") then
		Err_Msg = Err_Msg & "<li>你必须在网址前加上 <b>http://</b> or <b>https://</b></li>"
	end if
	if strAuthType = "nt" and ChkAccountReg = "true" then
		Err_Msg = Err_Msg & "<li>NT用户帐号已经注册</li>"
	end if

	if strUniqueEmail = "1" then

		'## Forum_SQL
		strSql = "SELECT M_EMAIL FROM " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " WHERE M_EMAIL = '" & Trim(Request.Form("Email")) &"'"

		set rs = my_Conn.Execute (strSql)

		if rs.BOF and rs.EOF then 
			'## Do Nothing
		else 
			Err_Msg = Err_Msg & "<li>此邮件地址已经有人使用，请重新输入</li>"
		end if
		rs.close
		set rs = nothing
	end if
	if Len(Request.Form("Sig")) > 255 then
		Err_Msg = Err_Msg & "<li>个人签名不能超过255个字符 "
		Err_Msg = Err_Msg & "现在共有 <b>" & Len(Request.Form("Sig")) & "</b> 个字符</li>"
	end if
	if Err_Msg = "" then
		if Request.Form("Homepage") <> "" and lcase(Request.Form("Homepage")) <> "http://" and lcase(Request.Form("Homepage")) <> "https://" and lcase(Request.Form("Homepage")) <> "file:///" then
			regHomepage = Request.Form("Homepage")
		else
			regHomepage = " "
		end if

		'## Forum_SQL
		strSql = "INSERT INTO " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & "(M_NAME"
		if strAuthType = "nt" then
			strSql = strSql & ", M_USERNAME"
		end if

		strSql = strSql & ", M_PASSWORD"
		strSql = strSql & ", M_EMAIL"
' ############# Hide Email Mod #############
		strSql = strSql & ", M_HIDE_EMAIL"
' ##########################################
		strSql = strSql & ", M_DATE"
		strSql = strSql & ", M_COUNTRY"
		strSql = strSql & ", M_SIG"
		strSql = strSql & ", M_YAHOO"
		strSql = strSql & ", M_ICQ"
		strSql = strSql & ", M_AIM"
		strSql = strSql & ", M_POSTS"
		strSql = strSql & ", M_HOMEPAGE"
		strSql = strSql & ", M_LASTHEREDATE"
		strSql = strSql & ", M_STATUS"
		strSql = strSql & ", M_FIRSTNAME" 
		strSql = strSql & ", M_LASTNAME"
		strsql = strsql & ", M_CITY"
		strsql = strsql & ", M_STATE"
		strsql = strsql & ", M_PHOTO_URL"
		strsql = strsql & ", M_LINK1" 
		strSql = strSql & ", M_LINK2"
		strSql = strsql & ", M_AGE"
		strSql = strSql & ", M_MARSTATUS"
		strSql = strsql & ", M_SEX"
		strSql = strSql & ", M_OCCUPATION" 
		strSql = strSql & ", M_BIO" 
		strSql = strSql & ", M_HOBBIES"
		strsql = strsql & ", M_LNEWS" 
		strSql = strSql & ", M_QUOTE"  
'########################### PM NEW MEMBER #####################
		strSql = strSql & ", M_PMEMAIL"
'########################### PM NEW MEMBER #####################
		strSql = strSql & ") "
		strSql = strSql & " VALUES ("
		If strAutoLogon = "1" then
			strSql = strSql & "'" & Session(strCookieURL & "strNTUserFullName") & "'"
		Else
			strSql = strSql & "'" & ChkString(Request.Form("Name"),"name") & "'"
		end if
		if strAuthType = "nt" then
			strSql = strSql & ", " & "'" & strDBNTUserName & "'"
		end if
		strSql = strSql & ", " & "'" & ChkString(Request.Form("Password"),"password") & "'"
		strSql = strSql & ", " & "'" & ChkString(Request.Form("Email"),"email") & "'"
' ################################ Hide Email Mod ##################################
		strSql = strSql & ", " & "'" & ChkString(Request.Form("HideMail"),"") & "'"
' ##################################################################################		
		strSql = strSql & ", " & "'" & DateToStr(strForumTimeAdjust) & "'"
		strSql = strSql & ", " & "'" & ChkString(Request.Form("Country"),"name") & "'"
		strSql = strSql & ", " & "'" & ChkString(Request.Form("Sig"),"message") & "'"
		strSql = strSql & ", " & "'" & ChkString(Request.Form("YAHOO"),"name") & "'"
		strSql = strSql & ", " & "'" & ChkString(Request.Form("ICQ"),"number") & "'"
		strSql = strSql & ", " & "'" & ChkString(Request.Form("AIM"),"name") & "'"
		strSql = strSql & ", " & "0"
		strSql = strSql & ", " & "'" & ChkString(regHomepage,"url") & "'"
		strSql = strSql & ", " & "'" & DateToStr(strForumTimeAdjust) & "'"
		strSql = strSql & ", " & "1"
		if strfullName = "1" then
			strSql = strSql & ", '" & ChkString(Request.Form("FirstName"),"") & "'" 
			strSql = strSql & ", '" & ChkString(Request.Form("LastName"),"") & "'"  
		else
			strSql = strSql & ", ''" 
			strSql = strSql & ", ''"  
		end if
		if strCity = "1" then
			strsql = strsql & ", '" & ChkString(Request.Form("City"),"") & "'"    
		else
			strsql = strsql & ", ''"
		end if
		if strState = "1" then
			strsql = strsql & ", '" & ChkString(Request.Form("State"),"") & "'" 
		else
			strsql = strsql & ", ''" 
		end if
		if strPicture = "1" then
			strsql = strsql & ", '" & ChkString(Request.Form("Photo_URL"),"") & "'"  
		else
			strsql = strsql & ", ''"  
		end if
		if strFavLinks = "1" then
			strsql = strsql & ", '" & ChkString(Request.Form("LINK1"),"") & "'"
			strSql = strSql & ", '" & ChkString(Request.Form("LINK2"),"") & "'"
		else
			strsql = strsql & ", ''"
			strSql = strSql & ", ''"  
		end if
		if strAge = "1" then
			strSql = strsql & ", '" & ChkString(Request.Form("Age"),"") & "'" 
		else
			strSql = strsql & ", ''" 
		end if
		if strMarStatus = "1" then
			strSql = strSql & ", '" & ChkString(Request.Form("MarStatus"),"") & "'"
		else
			strSql = strSql & ", ''"  
		end if
		if strSex = "1" then
			strSql = strsql & ", '" & ChkString(Request.Form("Sex"),"") & "'"
		else
			strSql = strSql & ", ''"  
		end if
		if strOccupation = "1" then
			strSql = strSql & ", '" & ChkString(Request.Form("Occupation"),"") & "'"
		else
			strSql = strSql & ", ''"  
		end if
		if strBio = "1" then
			strSql = strSql & ", '" & ChkString(Request.Form("Bio"),"message") & "'"
		else
			strSql = strSql & ", ''"  
		end if
		if strHobbies = "1" then
			strSql = strSql & ", '" & ChkString(Request.Form("Hobbies"),"message") & "'" 
		else
			strSql = strSql & ", ''" 
		end if
		if strLNews = "1" then
			strsql = strsql & ", '" & ChkString(Request.Form("LNews"),"message") & "'"
		else
			strSql = strSql & ", ''"  
		end if
		if strQuote = "1" then
			strSql = strSql & ", '" & ChkString(Request.Form("Quote"),"message") & "'"
		else 
			strSql = strSql & ", ''" 
		end if
'########################### PM NEW MEMBER #####################
			strSql = strsql & ", '1'" 
'########################### PM NEW MEMBER #####################
		strSql = strSql & ")"

		my_Conn.Execute (strSql)
		
		docount
		
		regHomepage = ""
		
'########################### PM NEW MEMBER #####################
if intNewMemPM = "1" then
	Dim objDict
	Set objDict = Server.CreateObject("Scripting.Dictionary")
	set objRec = my_Conn.execute("SELECT * FROM " & strTablePrefix & "MODS WHERE M_NAME = 'NewMemPM'")

	while not objRec.EOF 
	  objDict.Add objRec.Fields.Item("m_code").Value, objRec.Fields.Item("m_value").Value
	  objRec.moveNext
	wend 

	Admin = GetMemberID(objDict.Item("Admin"))
	Subject = objDict.Item("Subject")
	Message = objDict.Item("Message")

	strSql = "INSERT INTO " & strTablePrefix & "PM "
	strSql = strSql & "(M_SUBJECT"
	strSql = strSql & ", M_FROM"
	strSql = strSql & ", M_TO"
	strSql = strSql & ", M_SENT"
	strSql = strSql & ", M_MESSAGE"
	strSql = strSql & ", M_PMCOUNT"
	strSql = strSql & ", M_READ"
	strSql = strSql & ", M_MAIL"
	strSql = strSql & ", M_OUTBOX"
	strSql = strSql & ") "
	strSql = strSql & " VALUES ("
	strSql = strSql & "'" & Subject & "'"
	strSql = strSql & ", " & "'" & Admin & "'"
	strSql = strSql & ", " & "'" & GetMemberID( Request.Form("Name")) & "'"
	strSql = strSql & ", " & "'" & DateToStr(strForumTimeAdjust) & "'"
	strSql = strSql & ", " & "'" & Message & "'"
	strSql = strSql & ", " & "'1'"
	strSql = strSql & ", " & "'0'"
	strSql = strSql & ", " & "'0'"
	strSql = strSql & ", " & 1 & ")"
	
	my_Conn.Execute (strSql)
end if
'########################### PM NEW MEMBER #####################

'Rem User Field Code #######################################
			if (strUseExtendedProfile) then
				if (intUserFields > 0) then
					memID = getMemberID(Request.Form("Name"))
					UserstrSql = "SELECT * " 
					UserstrSql = UserstrSql & " FROM " & strTablePrefix & "USERFIELDS "
					set rsIDs = my_Conn.Execute(UserstrSql)
					do until rsIDs.EOF
						sInputName = rsIDs("USR_SHORTNAME")
							UpdSql = "INSERT INTO " & strMemberTablePrefix & "MEMBERFIELDS "
							UpdSql = UpdSql & "VALUES ( '" & memID & "','" & rsIDs("USR_FIELD_ID") & "','"
							UpdSql = UpdSql & Request.Form(sInputName) & "')"
							my_Conn.Execute(UpdSQL)
							UpdSQL = ""
						rsIDs.MoveNext
					loop
					rsIDs.Close
				end if
			end if
'Rem User Field Code #######################################

		if lcase(strEmail) = "1" then
			'## Emails Message to the Author of this Reply.  
			'## This needs to be Edited to use your own email component
			'## if you don't have one, try the w3Jmail component from www.dimac.net it's free!
			strRecipientsName = Request.Form("Name")
			strRecipients = Request.Form("Email")
			strFrom = strSender
			strFromName = strForumTitle
			strsubject = strForumTitle & "注册通知"
			strMessage = "Hello " & Request.Form("name") & vbCrLf & vbCrLf
			strMessage = strMessage & "这是一封来自 " & strForumTitle & " 的注册通知，恭喜你注册成功，现在你可以在 " & strForumURL & "发表新文章或回复文章并且可以使用我们所提供的资源" & vbCrLf & vbCrLf
			if strAuthType="db" then
				strMessage = strMessage & "你的密码是：" & Request.Form("Password") & vbCrLf & vbCrLf
			end if
			strMessage = strMessage & "你可以经由点击""资料""来修改你的资料" & vbCrLf & vbCrLf
			strMessage = strMessage & "赶快加入论坛吧！"
%>
<!--#INCLUDE FILE="inc_mail.asp" -->
<%
		end if
	else
		if (InStr(Request.ServerVariables("HTTP_REFERER"), "policy.asp") <> 0) or (Err_Msg <> "") then 
%>

	<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">你输入的资料有问题或没有填写完整</font></p>

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

<!--#INCLUDE FILE="inc_footer.asp" -->
<%		Response.End %>

<%			call ShowForm %>
<%'			On Error Resume next %>
<%			my_Conn.Close %>
<%			set my_Conn = nothing %>
<%			Response.End %>
<%		else %>
	<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">发现错误</font></p>
    <p align=center><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>>You must first read the disclaimer/privacy statement by clicking on &quot;<a href="policy.asp">register</a>&quot; link in the main forum window.</p>
    <p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript:history.go(-1)">返回修正问题。</a></font></p>
    <!--#INCLUDE FILE="inc_footer.asp" -->
<%			Response.End %>
<%		end if %>
<%	end if %>
	<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">你的注册已经完成！</font></p>
	<p align=center><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>>你现在可以发表文章<% if strAuthType="db" then %>使用你的新用户名和密码<% end if %>。</font></p>
<%
	if strAuthType = "db" then

		select case ChkUser2(Request.Form("Name"), Request.Form("Password"))
			case 1, 2, 3, 4
				Call DoCookies("false")
				strLoginStatus = 1
			case else
				strLoginStatus = 0
		end select
	end if
%>
<%	If strAutoLogon = 1 then
  		Response.Redirect "default.asp"
	Else %>
		<meta http-equiv="Refresh" content="0; URL=<% =Request.Form("refer")%>">
<%	End if %>
    <p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="<% =Request.Form("refer")%>">返回论坛</font></a></p>
<% end if %>
<%' On Error Resume next %>
<!--#INCLUDE FILE="inc_footer.asp" -->
<% 
sub DoCount

	'## Forum_SQL - Updates the Totals table by adding 1 to U_COUNT
	strSql = "UPDATE " & strTablePrefix & "TOTALS "
	strSql = strSql & " SET " & strTablePrefix & "TOTALS.U_COUNT = " & strTablePrefix & "TOTALS.U_COUNT + 1"

	my_Conn.Execute (strSql)

end sub
%>
<% sub ShowForm() %>
<form action="register.asp?mode=DoIt" method="post" id=form1 name=form1>
<input name="Refer" type="hidden" value="<% =Request.Form("Refer") %>">
<table width="400" border=0 align=center>
  <tr>
    <td><!--#INCLUDE FILE="inc_profile.asp" --></td>
  </tr>
</table>
</form>
<% end sub %>
