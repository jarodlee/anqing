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
<%
'#################################################################################
'## Variable declaration 
'#################################################################################
dim strSelecSize
dim intCols, intRows
%>
<!--#INCLUDE FILE="config.asp" --> 
<!--#INCLUDE FILE="inc_functions.asp" -->
<!--#INCLUDE FILE="inc_code.asp" -->
<%
'#################################################################################
'## Initialise variables 
'#################################################################################
strSelectSize = Request.Form("SelectSize") 
strRqMethod = Request.QueryString("method")
strCkName = Request.Cookies(strUniqueID & "User")("Name")
strCkPassWord = Request.Cookies(strUniqueID & "User")("Pword")
'#################################################################################
'## Page-code start
'#################################################################################

if strSelectSize = "" or IsNull(strSelectSize) then 
	strSelectSize = Request.Cookies(strCookieURL & "strSelectSize")
end if
if not(IsNull(strSelectSize)) then 
	Response.Cookies(strCookieURL & "strSelectSize") = strSelectSize
	Response.Cookies(strCookieURL & "strSelectSize").expires = Now() + 365
end if
%>
<!--#INCLUDE FILE="inc_top.asp" -->
<%
select case strSelectSize
	case "1"
		intCols = 45
		intRows = 6
	
	case "2"
		intCols = 80
		intRows = 12

	case "3"
		intCols = 90
		intRows = 12
	
	case "4"
		intCols = 130
		intRows = 15
	
	case else
		intCols = 45
		intRows = 6
		
end select

%>
<script language="JavaScript">
<!--
function autoReload(objform)
{
	objform.submit()
}

function OpenPreview()
{
	var curCookie = "strMessagePreview=" + escape(document.PostTopic.Message.value);
	document.cookie = curCookie;
	popupWin = window.open('pop_preview.asp', 'preview_page', 'scrollbars=yes,width=750,height=450')	
}


//-->
</script>
<% 


Msg = ""


if Request.Querystring("method") = "ReplyQuote" or Request.Querystring("method") = "Forward" then
	
	'## Forum_SQL
	strSql = "SELECT * "
	strSql = strSql & " FROM " & strTablePrefix & "PM "
	strSql = strSql & " WHERE " & strTablePrefix & "PM.M_ID = " & Request.QueryString("id")

	set rs = my_Conn.Execute (strSql)
	
	strAuthor = rs("M_FROM") '**

	if Request.QueryString("method") = "Edit" then
		TxtMsg = rs("M_MESSAGE")
	else
		if Request.Querystring("method") = "ReplyQuote" then
		
			TxtMsg = "[quote]" & vbCrLf
			TxtMsg = TxtMsg & rs("M_MESSAGE") & vbCrLf
			TxtMsg = TxtMsg & "[/quote]"
		end if
		if Request.Querystring("method") = "Forward" then
			TxtMsg = "----- 转发信息 -----" & vbCrLf
			TxtMsg = TxtMsg & rs("M_MESSAGE") & vbCrLf
			TxtMsg = TxtMsg & "-----------------------------"
		end if
		
	end if
	'boolReply = rs("R_MAIL") 
end if


%>


<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
<% =Msg %>
</font></p>

<form action="privatesend_info.asp" method="post" name="PostTopic">
<input name="R_SUBJECT" type="hidden" value="<% =Request.QueryString("R_SUBJECT") %>">
<input name="Method_Type" type="hidden" value="<% =Request.QueryString("method") %>">
<input name="Type" type="hidden" value="<% =Request.QueryString("type") %>">
<input name="REPLY_ID" type="hidden" value="<% =Request.QueryString("REPLY_ID") %>">
<input name="TOPIC_ID" type="hidden" value="<% =Request.QueryString("TOPIC_ID") %>">
<input name="Author" type="hidden" value="<% =strAuthor %>">
<input name="M" type="hidden" value="<% =Request.QueryString("M") %>">
<input name="Refer" type="hidden" value="<% =Request.ServerVariables("HTTP_REFERER") %>">
<input name="cookies" type="hidden" value="yes">

<table border="0" cellspacing="1" cellpadding="0" align=center>
  <tr>
    <td bgcolor="<% =strPopUpBorderColor %>">
    <table border="0" cellspacing="1" cellpadding="4">
<% 
if mlev = 4 or _
mlev = 3 or _
mlev = 2 or _
mlev = 1 then 
%>
    <input name="UserName" type="hidden" Value="<% =Request.Cookies(strUniqueID & "User")("Name")%>">
    <input name="Password" type="hidden" value="<% =Request.Cookies(strUniqueID & "User")("PWord")%>">
<% else %>
<% 
	if (lcase(strNoCookies) = "1") or _
	(Request.Cookies(strUniqueID & "User")("Name") = "" or _
	Request.Cookies(strUniqueID & "User")("PWord") = "") then 
%>
      <tr>
        <td bgColor="<% =strPopUpTableColor %>" noWrap vAlign="top" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">用户名称：</font></td>
        <td bgColor="<% =strPopUpTableColor %>"><input name="UserName" maxLength="25" size="25" type="text" value=""></td>
      </tr>
      <tr>
        <td bgColor="<% =strPopUpTableColor %>" noWrap vAlign="top" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">你的密码：</font></td>
        <td bgColor="<% =strPopUpTableColor %>" valign="top"><input name="Password" maxLength="13" size="13" type="password" value=""></td>
      </tr>
<%	end if %>
<% end if %>



<% 
if Request.QueryString("method") = "Topic" then 
%>
      <tr>
        <td bgColor="<% =strPopUpTableColor %>" noWrap vAlign="top" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">主题：</font></td>
        <td bgColor="<% =strPopUpTableColor %>"><input maxLength="50" name="Subject" value="<% =Trim(ChkString(TxtSub,"display")) %>" size="50"></td>
      </tr>
<% end if 
if Request.QueryString("method") = "Topic" or Request.Querystring("method") = "Forward" then %>
<script language="JavaScript">
<!-- hide from JavaScript-challenged browsers
function pmmembers() { var MainWindow = window.open ("pm_pop_members.asp", "","toolbar=no,location=no,menubar=no,scrollbars=yes,width=250,height=500,top=100,left=100,resizeable=yes,status=no");
}
// done hiding -->
</script>
      <tr>
        <td bgColor="<% =strPopUpTableColor %>" noWrap vAlign="top" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">收件人：</font></td>
        <td bgColor="<% =strPopUpTableColor %>"><input maxLength="50" name="sendto" value="<% =Request.Querystring("mname")%>" size="50">&nbsp;&nbsp;<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>-1"><a href="JavaScript:pmmembers();">会员列表</a></font></td>
      </tr>
<% end if %>

<% 
if Request.QueryString("method") = "Reply" or _
Request.QueryString("method") = "ReplyQuote" or _
Request.Querystring("method") = "Forward" or _
Request.QueryString("method") = "Topic" then 
%>
<% if strAllowForumCode = 1 then %>
<!--#INCLUDE FILE="inc_post_buttons.asp" -->
	  	<tr>
		<td bgColor="<% =strPopUpTableColor %>"  valign=top align="right">
        		<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>模式：</b>
        	</td>

		<td bgColor="<% =strPopUpTableColor %>" align="left" valign="top">
	      		<select name="font" onChange="thelp(this.options[this.selectedIndex].value)">
			<option value="1">帮助&nbsp;</option>
			<option selected value="2">完全&nbsp;</option>
			<option value="0">基本&nbsp;</option>
	      		</select>          
	    		</font><a href="JavaScript:openWindow3('pop_help.asp#mode')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
		</td>
          </tr> 
<% end if %>
      <tr>
        <td bgColor="<% =strPopUpTableColor %>" noWrap vAlign="top" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">内容：<br>
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
            * 本版专用代码开启<br>
<% else %>
            * 本版专用代关闭<br>
<% end if %>
            </font></td>
          </tr>
        </table>
        </font></td>
        <td bgColor="<% =strPopUpTableColor %>"><textarea cols="45" name="Message" rows="8" wrap="VIRTUAL"><% =Trim(CleanCode(TxtMsg))%></textarea></td>
      </tr>
<% end if %>
<% 
if Request.QueryString("method") = "Reply" or _
Request.QueryString("method") = "ReplyQuote" or _
Request.Querystring("method") = "Forward" or _
Request.QueryString("method") = "Topic"  then 
%>
      <tr>
        <td bgColor="<% =strPopUpTableColor %>">&nbsp;</td>
        <td bgColor="<% =strPopUpTableColor %>">
<% 
	if Request.QueryString("method") = "Reply" or _
	Request.QueryString("method") = "ReplyQuote" or _
	Request.Querystring("method") = "Forward" or _
	Request.QueryString("method") = "Topic" then 
%>
        <table border=0>
          <tr>
<%		if lcase(stricons) = "1" then %>
            <td><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
          
<%		end if %>

            </td>
          </tr>
        </table>
<%	end if %>
<% 
	if Request.QueryString("method") = "Edit" or _
	Request.QueryString("method") = "EditTopic" or _
	Request.QueryString("method") = "Reply" or _
	Request.QueryString("method") = "ReplyQuote" or _
	Request.Querystring("method") = "Forward" or _
	Request.QueryString("method") = "Topic" or _
	Request.QueryString("method") = "TopicQuote" then 
%>
<%
		if Request.QueryString("method") = "Reply" or _
		
		Request.QueryString("method") = "ReplyQuote" or _
		Request.Querystring("method") = "Forward" or _
		Request.QueryString("method") = "Topic" or _
		Request.QueryString("method") = "TopicQuote" then 
%>
<%'		if rsMember("M_SIG") <> " " then %>
        <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
        <input name="Sig" type="checkbox" value="yes" checked<% =Chked(Request.Cookies("User")("sig"))%>>包含个人签名<br>
        </font>
<%'		end if %>
<%		end if %>
<%		if lcase(strEmail) = "1" then %>
<%			if Request.QueryString("method") = "Topic" or _ 
			Request.QueryString("method") = "EditTopic" then %>  
        <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
        
        </font>
<%			else %>
<%
				if Request.QueryString("method") = "Reply" or _
				Request.QueryString("method") = "Edit" or _ 
				Request.QueryString("method") = "ReplyQuote" or _
				Request.Querystring("method") = "Forward" or _
				Request.QueryString("method") = "TopicQuote" then 
		end if %>
<%			end if %>
<%		end if %>
<%	end if %>
<% end if %>
        </td>
      </tr>
      <tr>
        <td bgColor="<% =strPopUpTableColor %>">&nbsp;</td>
        <td bgColor="<% =strPopUpTableColor %>"><input name="Submit" type="submit" value="提交信息">&nbsp;<input name="Reset" type="reset" value="全部清除"></td>
      </tr>


    </table>
    </td>
  </tr>
</table>
</form>

<%
If Request.QueryString("Method") = "Reply" or _
Request.QueryString("Method") = "TopicQuote" or _
Request.Querystring("method") = "Forward" or _
Request.QueryString("Method") = "ReplyQuote" then
%>

    <center>
    <table bgcolor="<% =strHeadCellColor %>" border="0" width="95%" cellspacing="1" cellpadding="4">
      <tr>
        <td bgcolor="<% =strHeadCellColor %>" colspan="2" align="center"><font <% =strDefaultFontFace %> size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">原  &nbsp;&nbsp;&nbsp; 讯  &nbsp;&nbsp;&nbsp; 息</font></td>
      </tr>
<%
	'## Forum_SQL
	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strTablePrefix & "PM.M_MESSAGE " 
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS, " & strTablePrefix & "PM "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "PM.M_FROM AND "
	strSql = strSql & "       " & strTablePrefix & "PM.M_ID = " &  Request.QueryString("ID")

	set rs = my_Conn.Execute (strSql) 

	Response.Write "      <tr>" & vbCrLf
	Response.Write "        <td bgcolor='" & strForumFirstCellColor & "' valign=top width='" & strTopicWidthLeft & "'"
	if lcase(strTopicNoWrapLeft) = "1" then
		Response.Write " nowrap"
	end if 
	Response.Write "><font color='" & strForumFontColor & "' face='" & strDefaultFontFace & "' size='2'><b>" & ChkString(rs("M_NAME"),"display") & "</b></font></td>" & vbCrLf
	Response.Write "        <td bgcolor='" & strForumCellColor & "' valign='top' width='" & strTopicWidthRight & "'" 
	if lcase(strTopicNoWrapRight) = "1" then
		Response.Write " nowrap"
	end if 
	Response.Write "><font color='" & strForumFontColor & "' face='" & strDefaultFontFace & "' size='2'>" & formatStr(rs("M_MESSAGE")) & "</font></td>" & vbCrLf
	Response.Write "      </tr>" & vbCrLf
	Response.Write "    </TABLE>" & vbCrLf
	
	Response.Write "</FONT>" & vbCrLf

end if
%> 

  	  <p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="default.asp">返回论坛首页</font></a></p>
<!--#INCLUDE FILE="inc_footer.asp" -->