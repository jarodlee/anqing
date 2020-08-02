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
<!--#INCLUDE FILE="inc_top_short.asp" -->
<% 
set rs = Server.CreateObject("ADODB.RecordSet")

'## Forum_SQL
strSql = "SELECT * FROM " & strMemberTablePrefix & "MEMBERS "
strSql = strSql & " WHERE MEMBER_ID = " & Request.QueryString("ID")

rs = my_Conn.Execute (strSql)
Set hideMail = rs("M_HIDE_EMAIL")
if (hideMail = 1) then
%>
<p>抱歉！该会员不能收到来自论坛的电子邮件！<p>
<p> </p>

<%else%>



<p><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">发送一封电子邮件</font></p>
<%
if lcase(strEmail) = "1" then
	Err_Msg = ""
	if Request.Form("YName") = "" then 
		Err_Msg = Err_Msg & "<li>你必须输入你的用户名</li>"
	end if
	if Request.Form("YEmail") = "" then 
		Err_Msg = Err_Msg & "<li>你必须输入你的电子邮件地址！</li>"
	else
		if EmailField(Request.Form("YEmail")) = 0 then 
			Err_Msg = Err_Msg & "<li>必须正确填写你的email地址</li>"
		end if
	end if
	if Request.Form("Name") = "" then 
		Err_Msg = Err_Msg & "<li>你必须输入收件者姓名！</li>"
	end if
	if Request.Form("Email") = "" then 
		Err_Msg = Err_Msg & "<li>你必须输入收件人电子邮件地址！</li>"
	else
		if EmailField(Request.Form("Email")) = 0 then 
			Err_Msg = Err_Msg & "<li>你必须正确填写收件人电子邮件地址！</li>"
		end if
	end if
	if Request.Form("Msg") = "" then 
		Err_Msg = Err_Msg & "<li>你必须输入邮件内容！</li>"
	end if
	'##  Emails Message to the Author of this Reply.  
	'##  This needs to be Edited to use your own email component
	'##  if you don't have one, try the w3Jmail component from www.dimac.net it's free!
	if Request.QueryString("mode") = "DoIt" then
		if (Err_Msg = "") then
			strRecipientsName = Request.Form("Name")
			strRecipients = Request.Form("Email")
			strFrom = Request.Form("YEmail")
			strFromName = Request.Form("YName")
			strSubject = "来自 " & strForumTitle & " 的 " & Request.Form("YName")
			strMessage = "你好： " & Request.Form("Name") & vbCrLf & vbCrLf
			strMessage = strMessage & "这封邮件来自 : " & Request.Form("YName") & " (" & Request.Form("YEmail") & ") " & vbCrLf & vbCrLf 
			strMessage = strMessage & "我们的网站地址: " & strForumURL & vbCrLf & vbCrLf
			strMessage = strMessage & Request.Form("Msg") & vbCrLf & vbCrLf

			if strFrom <> "" then 
				strSender = strFrom
			end if
%>
<!--#INCLUDE FILE="inc_mail.asp" -->

			<p><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">电子邮件已寄出！</font></p>
				
<%
		else
%>
	<p><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">这封电子邮件有问题，可能有输入错误或没有填写完整。</font></p>

	<table>
	  <tr>
	    <td><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
		<ul>
		<% =Err_Msg %>
		</ul></font>
	    </td>
	  </tr>
	</table>

	<p><font size="<% =strDefaultFontSize %>"><a href="JavaScript:history.go(-1)">请返回重新输入</a></font></p>

	<p><font size="<% =strDefaultFontSize %>"><a href="JavaScript:onClick= window.close()">关闭窗口</A></font></p>
<%
			Response.End 
		end if
	else 
		Err_Msg = ""
		if rs("M_EMAIL") <> " " then
				strSql =  "SELECT M_NAME, M_USERNAME, M_EMAIL "
				strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS"
				strSql = strSql & " WHERE "&Strdbntsqlname&" = '" & STRdbntUserName & "'"

			set rs2 = my_conn.Execute (strSql)
			YName = ""
			YEmail = ""

			if (rs2.EOF or rs2.BOF)  then
				if strLogonForMail <> "0" then 
		
					Err_Msg = Err_Msg & "<li>发送电子邮件前必须先登陆论坛！</li>"
%>
					<table>
					 <tr>
					  <td><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
					  <ul>
						<% =Err_Msg %>
					  </ul></font>
					  </td>
					 </tr>
					</table>
<%					If InStr(Request.ServerVariables("HTTP_REFERER"), "pop_profile.asp") <> 0 then %>
						<p><font size="<% =strDefaultFontSize %>"><a href="JavaScript:history.go(-1)">回到 <% =rs("M_NAME")%> 的会员资料</a></font></p>
<%					end if %>

					<p><font size="<% =strDefaultFontSize %>"><a href="JavaScript:onClick= window.close()">关闭窗口</A></font></p>
<%
					Response.End
				end if
			else
				YName = Trim("" & rs2("M_NAME"))
				YEmail = Trim("" & rs2("M_EMAIL"))
			end if
			rs2.close
			set rs2 = nothing
%>

<form action="pop_mail.asp?mode=DoIt&id=<% =Request.QueryString("ID") %>" method=Post id=Form1 name=Form1>
<input type=hidden name="Page" value="<% =Request.QueryString("page") %>">

<table border="0" width="90%" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="<% =strPopUpBorderColor %>">
    <table border="0" width="100%" cellspacing="1" cellpadding="1">
      <tr>
        <td bgColor=<% =strPopUpTableColor %> align="right" nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">收件人姓名：</FONT></b></td>
        <td bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% =rs("M_NAME") %><input type=hidden name="Name" value="<% =rs("M_NAME") %>"></font></td>
      </tr>
      <tr>
        <td bgColor=<% =strPopUpTableColor %> align="right" nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">收件人Email地址：</FONT></b></td>
        <td bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% =rs("M_EMAIL") %><input type=hidden name="Email" value="<% =rs("M_EMAIL") %>"></font></td>
      </tr>                
      <tr>
        <td bgColor=<% =strPopUpTableColor %> align="right" nowrap><b><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>>发件人姓名：</font></b></td>
        <td bgColor=<% =strPopUpTableColor %>><input name=YName type=<% if YName <> "" then Response.Write("hidden") else Response.Write("text") end if %> value="<% = YName %>" size=25><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>> <% if YName <> "" then Response.Write(YName) end if %></font></td>
      </tr>
      <tr>
        <td bgColor=<% =strPopUpTableColor %> align="right" nowrap><b><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>>发件人Email地址：</font></b></td>
        <td bgColor=<% =strPopUpTableColor %>><input name=YEmail type=<% if YEmail <> "" then Response.Write("hidden") else Response.Write("text") end if %> value="<% = YEmail %>" size=25><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>> <% if YEmail <> "" then Response.Write(YEmail) end if %></font></td>
      </tr> 
      <tr>
        <td bgColor=<% =strPopUpTableColor %> colspan=2><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">信件内容：</FONT></b></td>
      </tr>
      <tr>
        <td bgColor=<% =strPopUpTableColor %> colspan=2><textarea name="Msg" cols=40 rows=5></textarea></td>
      </tr>                    
      <tr>
        <td bgColor=<% =strPopUpTableColor %> colspan=2 align=center><Input type=Submit value="发送" id=Submit1 name=Submit1></TD>
      </tr>
    </table>
    </td>
  </tr>
</table>
</form>
<%		else %>
<p>该用户并未登录其电子邮箱</p>
<%		end if %>
<%	end if %>
<% else %>
<p>点击发送送一封电子邮件给 <a href="mailto:<% =rs("M_EMAIL")%>"><% =rs("M_NAME")%></a></p>
<% end if %>
<% If InStr(Request.ServerVariables("HTTP_REFERER"), "pop_profile.asp") <> 0 then %>
<p><font size="<% =strDefaultFontSize %>"><a href="JavaScript:history.go(-1)">回到 <% =rs("M_NAME")%> 的会员资料</a></font></p>
<% end if 
end if%>

<% set rs = nothing %>
<!--#INCLUDE FILE="inc_footer_short.asp" -->
