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
<script language="JavaScript">
 function changeTopics(){
 
	document.Form2.Tcount.value = document.Form1.TopicCount.value
	document.Form2.submit()
}
</script> 

<%  if Request.QueryString("mode") <> "DoIt" then
	if Request.QueryString("TOPIC_ID") = "" then
		TopicID = Request("hTopicID")
	else
		TopicID = Request.QueryString("TOPIC_ID")
	end if
	'## Forum_SQL - Get Origional Posting
	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strTablePrefix & "TOPICS.T_DATE, " & strTablePrefix & "TOPICS.T_SUBJECT, " & strTablePrefix & "TOPICS.T_AUTHOR, " & strTablePrefix & "TOPICS.TOPIC_ID, " & strTablePrefix & "TOPICS.T_MESSAGE "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS, " & strTablePrefix & "TOPICS "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "TOPICS.T_AUTHOR "
	strSql = strSql & " AND   " & strTablePrefix & "TOPICS.TOPIC_ID = " &  TopicID 
	set rs4 = my_Conn.Execute (strSql)
	'## Forum_SQL - Get all topicsFrom DB
	strSql ="SELECT " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strTablePrefix & "REPLY.REPLY_ID, " & strTablePrefix & "REPLY.R_AUTHOR, " & strTablePrefix & "REPLY.TOPIC_ID, " & strTablePrefix & "REPLY.R_MESSAGE, " & strTablePrefix & "REPLY.R_DATE "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS, " & strTablePrefix & "REPLY "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "REPLY.R_AUTHOR "
	strSql = strSql & " AND   TOPIC_ID = " & TopicID & " "
	strSql = strSql & " ORDER BY " & strTablePrefix & "REPLY.R_DATE"
	set rs3 = Server.CreateObject("ADODB.Recordset")
	rs3.cachesize = 20
	rs3.open  strSql, my_Conn, 3
end if
		
if Request("Tcount") = ""	Then 
	TopicCount = 9999
else
	TopicCount = Request("Tcount")
end if
if Request.QueryString("mode") = "DoIt" then
	Err_Msg = ""
	if (Request.Form("YName") = "") then 
		Err_Msg = Err_Msg & "<li>你必须输入你的姓名！</li>"
	end if
	if (Request.Form("YEmail") = "") then 
		Err_Msg = Err_Msg & "<li>你必须输入你的邮件地址！</li>"
	else
		if (EmailField(Request.Form("YEmail")) = 0) then 
			Err_Msg = Err_Msg & "<li>必须填写正确的邮件地址</li>"
		end if
	end if
	if (Request.Form("Name") = "") then 
		Err_Msg = Err_Msg & "<li>你必须输入收件人姓名！</li>"
	end if
	if (Request.Form("Email") = "") then 
		Err_Msg = Err_Msg & "<li>你必须输入收件人邮件地址！</li>"
	else
		if (EmailField(Request.Form("Email")) = 0) then 
			Err_Msg = Err_Msg & "<li>你必须输入正确的收件人邮件地址！</li>"
		end if
	end if
	if (Request.Form("Msg") = "") then 
		Err_Msg = Err_Msg & "<li>你必须输入邮件内容！</li>"
	end if
	'##  Emails Topic To a Friend.  
	'##  This needs to be Edited to use your own email component
	'##  if you don't have one, try the w3Jmail component from www.dimac.net it's free!
	if lcase(strEmail) = "1" then
		if (Err_Msg = "") then
			strRecipientsName = Request.Form("Name")
			strRecipients = Request.Form("Email")
			strSubject = "From: " & Request.Form("YName") & " - Interesting Discussion at " & strForumTitle
			strMessage = "Hello " & Request.Form("Name") & vbCrLf & vbCrLf
			strMessage = strMessage & Request.Form("Msg") & vbCrLf & vbCrLf
			strMessage = strMessage & "You received this e-mail from: " & Request.Form("YName") & " " & Request.Form("YEmail")
%>
<!--#INCLUDE FILE="inc_mail.asp" -->
<%
			Response.Write "<p><font face='" & strDefaultFontFace & "' size=" & strHeaderFontSize & ">邮件已寄出！</font></p>" & vbCrLf
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
	end if
else 
%>


<% 


		'## Forum_SQL
		strSql =  "SELECT M_NAME, M_USERNAME, M_EMAIL "
		strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS"
		strSql = strSql & " WHERE "&Strdbntsqlname&" = '" & STRdbntUserName & "'"

		set rs = my_conn.Execute (strSql)
		YName = ""
		YEmail = ""

		if (rs.EOF or rs.BOF)  then
			if strLogonForMail <> "0" then 
			
				Err_Msg = Err_Msg & "<li>发送电子邮件前必须先登陆论坛！</li>"
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

			<p><font size="<% =strDefaultFontSize %>"><a href="JavaScript:onClick= window.close()">关闭窗口</A></font></p>
<%
				Response.End
			end if
		else
			YName = Trim("" & rs("M_NAME"))
			YEmail = Trim("" & rs("M_EMAIL"))
		end if
		rs.close
		set rs = nothing

%>

<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">将本主题推荐给你朋友</font>
<form action="pop_send_to_friend.asp?<% =Request.QueryString %>" name=Form2 id=Form2><input type=hidden name="Tcount" value="">
<input type=hidden name="hTopicID" value="<%= TopicID %>">
</form>
<form action="pop_send_to_friend.asp?mode=DoIt" method=post id=Form1 name=Form1>
<input type=hidden name="Page" value="<% =Request.QueryString %>">
	
<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="<% =strPopUpBorderColor %>">
    <table border="0" width="100%" cellspacing="1" cellpadding="2">
      <TR>
        <TD bgColor=<% =strPopUpTableColor %> align="right" nowrap><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>>收件人姓名：</font></td>
        <TD bgColor=<% =strPopUpTableColor %>><input type=text name="Name" size=25></td>
      </tr>
      <TR>
        <TD bgColor=<% =strPopUpTableColor %> align="right" nowrap><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>>收件人邮件地址：</font></td>
        <TD bgColor=<% =strPopUpTableColor %>><input type=text name="Email" size=25></td>
      </tr>                
      <tr>
        <td bgColor=<% =strPopUpTableColor %> align="right" nowrap><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>>发件人姓名：</font></td>
        <td bgColor=<% =strPopUpTableColor %>><input name=YName type=<% if YName <> "" then Response.Write("hidden") else Response.Write("text") end if %> value="<% = YName %>" size=25><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>> <% if YName <> "" then Response.Write(YName) end if %></font></td>
      </tr>
      <tr>
        <td bgColor=<% =strPopUpTableColor %> align="right" nowrap><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>>发件人邮件地址：</font></td>
        <td bgColor=<% =strPopUpTableColor %>><input name=YEmail type=<% if YEmail <> "" then Response.Write("hidden") else Response.Write("text") end if %> value="<% = YEmail %>" size=25><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>> <% if YEmail <> "" then Response.Write(YEmail) end if %></font></td>
      </tr> 

      <tr>
        <td bgColor=<% =strPopUpTableColor %> align="right" nowrap><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>>信件内容：</font></td>
		<td bgColor=<% =strPopUpTableColor %> nowrap><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>>要包含几篇回复？</font><input type="text" name="TopicCount" value="5" size="4"> <input type="button" value="变更" name="change" onclick="changeTopics()"></td>
      </tr>
      <tr>
        <td bgColor=<% =strPopUpTableColor %> colspan=2 align=center><textarea name="Msg" cols="60" rows=10 readonly>在 <% =strForumTitle %> (<% =strForumURL%>) 里，我找到一篇你或许会感兴趣的文章：<% =vbCrLf & strForumURL%>/link.asp?TOPIC_ID=<%= TopicID & vbCrLf & vbCrLf %>这是那篇文章中的一部份：<% =vbCrLf %>
==========================================================

主题：<% =rs4("T_Subject")%>
作者：<% =rs4("M_NAME")%>
时间：<% =ChkDate(rs4("T_DATE")) %>
内容：<% =vbCrLf & rs4("T_MESSAGE") %>
<% i = 0
	rec=0 
	If rs3.EOF or rs3.BOF then  '## No categories found in DB
		Response.Write ""
	Else
		rs3.movefirst
		do until (rs3.EOF) or (rec > cInt(TopicCount))%>
----------------------------------------------------------
作者：<% =rs3("M_NAME")%>
时间：<% =ChkDate(rs3("R_DATE")) %>
内容：<% =vbCrLf & rs3("R_MESSAGE") %>
<% 
rs3.MoveNext
		    strI = strI + 1
		    if strI = 2 then 
				strI = 0
			end if
		    rec = rec + 1
		loop
	end if
 %>
==========================================================
￣讯息到此结束￣
<br>
欢迎加入我们的论坛：<% =strForumTitle%>
<% =strForumURL%>
</textarea></td>
      </tr>                    
      <tr>
        <td bgColor=<% =strPopUpTableColor %> colspan=2 align=center><input type=submit value="提交" id=Submit1 name=Submit1></td>
      </tr>
    </table>
    </td>
  </tr>
</table>
</form>
<% end if %>
<!--#INCLUDE FILE="inc_footer_short.asp" -->
