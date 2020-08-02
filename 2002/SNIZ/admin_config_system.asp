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
<!--#INCLUDE file="inc_top.asp" -->
<table border="0" width="100%">
  <tr>
    <td width="33%" align="left" nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    <img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="default.asp">返回论坛首页</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="admin_home.asp">论坛管理中心</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" height=15 width=15 border="0">&nbsp;论坛基本设置<br>
    </font></td>
  </tr>
</table>
<%
	if Request.Form("Method_Type") = "Write_Configuration" then 
		Err_Msg = ""
		if Request.Form("strTitleImage") = "" then 
			Err_Msg = Err_Msg & "<li>你必须输入论坛logo地址</li>"
		end if
		if Request.Form("strHomeURL") = "" then 
			Err_Msg = Err_Msg & "<li>你必须输入网站首页（相对路径或完整连结）</li>"
		end if
		if Request.Form("strForumURL") = "" then 
			Err_Msg = Err_Msg & "<li>你必须输入论坛地址</li>"
		end if
		if (left(lcase(Request.Form("strForumURL")), 7) <> "http://" and left(lcase(Request.Form("strForumURL")), 8) <> "https://") and Request.Form("strHomeURL") <> "" then
			Err_Msg = Err_Msg & "<li>你必须在连接前加上 <b>http://</b>, <b>https://</b> or <b>file://</b></li>"
		end if
		if (right(lcase(Request.Form("strForumURL")), 1) <> "/") then
			Err_Msg = Err_Msg & "<li>连接地址必须以 <b>/</b> 为结尾</li>"
		end if
		if Request.Form("strAuthType") <> strAuthType and strAuthType = "db" then 
			mLev = cint(ChkUser2(Request.Cookies(strUniqueID & "User")("Name"), Request.Cookies(strUniqueID & "User")("Pword")))
						
			if not(mLev = 4 and getMemberNumber(Request.Cookies(strUniqueID & "User")("Name")) = 1) then
				Err_Msg = Err_Msg & "<li>只有管理员才能修改论坛的授权模式</li>"
			else
				call NTauthenticate()
				if Session(strCookieURL & "userid") = "" then
					Err_Msg = Err_Msg & "<li>使用论坛前你必须先启动服务器的非匿名存取模式</li>"
				else	
					strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
					strSql = strSql & " SET " & strMemberTablePrefix & "MEMBERS.M_USERNAME = '" & Session(strCookieURL & "userid") & "'"
					strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & Request.Cookies(strUniqueID & "User")("Name") & "'; "

					my_Conn.Execute(strSql)			
					call NTauthenticate()
					call NTUser()	
				end if
			end if
		end if
		if (Request.Form("strAuthType") <> strAuthType) and strAuthType = "nt" then 
			mLev = cint(ChkUser2(Request.Cookies(strUniqueID & "User")("Name"), Request.Cookies(strUniqueID & "User")("Pword")))
			if not(mLev = 4 and getMemberNumber(Request.Cookies(strUniqueID & "User")("Name")) = 1) then
				Err_Msg = Err_Msg & "<li>只有管理员才能修改论坛的授权模式</li>"
			else
				Session(strCookieURL & "Approval") = "" 
			end if	
		end if
		if Err_Msg = "" then

			

			'## Forum_SQL
			strSql = "UPDATE " & strTablePrefix & "CONFIG "
			strSql = strSql & " SET C_STRFORUMTITLE = '" & ChkString(Request.Form("strForumTitle"),"title") & "'"
			strSql = strSql & ",    C_STRCOPYRIGHT = '" & ChkString(Request.Form("strCopyright"),"title") & "'"
			strSql = strSql & ",    C_STRTITLEIMAGE = '" & ChkString(Request.Form("strTitleImage"),"url") & "'"
			strSql = strSql & ",    C_STRHOMEURL = '" & ChkString(Request.Form("strHomeURL"),"url") & "'"
			strSql = strSql & ",    C_STRFORUMURL = '" & ChkString(Request.Form("strForumURL"),"url") & "'"
			strSql = strSql & ",    C_STRAUTHTYPE = '" & Request.Form("strAuthType") & "'"
			strSql = strSql & ",    C_STRSETCOOKIETOFORUM = " & Request.Form("strSetCookieToForum")
			strSql = strSql & ",    C_STRGFXBUTTONS = " & Request.Form("strGfxButtons")
			strSql = strSql & ",    C_STRSHOWIMAGEPOWEREDBY = " & Request.Form("strShowImagePoweredBy")
			strSql = strSql & " WHERE CONFIG_ID = " & 1
			my_Conn.Execute (strSql)
			Application(strCookieURL & "ConfigLoaded") = ""

%>
<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">论坛基本设置完毕</font></p>
<meta http-equiv="Refresh" content="0; URL=admin_home.asp">

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">修改完成！</font></p>

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="admin_home.asp">回到管理中心</font></a></p>
<%		else %>
	<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">你输入的资料有问题或没有完全填完</font></p>

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
<form action="admin_config_system.asp" method="post" id="Form1" name="Form1">
<input type="hidden" name="Method_Type" value="Write_Configuration">
<table border="0" cellspacing="1" cellpadding="4" align=center bgcolor="<% =strPopUpBorderColor %>">
  <tr valign="top">
    <td bgcolor="<% =strHeadCellColor %>" colspan="2"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>" color="<% =strHeadFontColor %>">主要设定</font></td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">论坛名称：</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    <input type="text" name="strForumTitle" size="25" value="<% if strForumTitle <> "" then Response.Write(strForumTitle) else '## do nothing %>">
    <a href="JavaScript:openWindow3('pop_config_help.asp#forumtitle')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">论坛版权说明：</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    <input type="text" name="strCopyright" size="25" value="<% if strCopyright <> "" then Response.Write(strCopyright) else Response.Write("&copy; 2000 Snitz Communications") %>">
    <a href="JavaScript:openWindow3('pop_config_help.asp#copyright')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">论坛logo图片地址：</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    <input type="text" name="strTitleImage" size="25" value="<% if strTitleImage <> "" then Response.Write(strTitleImage) else Response.Write("logo_snitz_forums_2000.gif") %>">
    <a href="JavaScript:openWindow3('pop_config_help.asp#titleimage')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">网站首页：</font></td></td>
    <td bgColor="<% =strPopUpTableColor %>">
    <input type="text" name="strHomeURL" size="25" value="<% if strHomeURL <> "" then Response.Write(strHomeURL) else '## Do Nothing %>">
    <a href="JavaScript:openWindow3('pop_config_help.asp#homeurl')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">论坛地址：</font></td></td>
    <td bgColor="<% =strPopUpTableColor %>">
    <input type="text" name="strForumURL" size="25" value="<% if strForumURL <> "" then Response.Write(strForumURL) else '## do nothing %>">
    <a href="JavaScript:openWindow3('pop_config_help.asp#forumurl')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">版本讯息：</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% if strVersion <> "" then Response.Write("[<b>"& strVersion & "</b>]") else Response.Write("<b>[Couldn't read version info..]</b>") %></font>
    <!--<a href="JavaScript:openWindow3('pop_config_help.asp#forumtitle')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>-->
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">授权模式：</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    DB: <input type="radio" class="radio" name="strAuthType" value="db" <% if strAuthType = "db" then Response.Write("checked") %>> 
    NT: <input type="radio" class="radio" name="strAuthType" value="nt" <% if strAuthType = "nt" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#AuthType')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </font></td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">设定 Cookie 到：</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    讨论区：<input type="radio" class="radio" name="strSetCookieToForum" value="1" <% if strSetCookieToForum = "1" then Response.Write("checked") %>> 
    网站：<input type="radio" class="radio" name="strSetCookieToForum" value="0" <% if strSetCookieToForum = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#SetCookieToForum')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </font></td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">使用图形按钮：</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    开启：<input type="radio" class="radio" name="strGfxButtons" value="1" <% if strGfxButtons = "1" then Response.Write("checked") %>> 
    关闭：<input type="radio" class="radio" name="strGfxButtons" value="0" <% if strGfxButtons = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#GfxButtons')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </font></td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">版权显示方式：</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    图形：<input type="radio" class="radio" name="strShowImagePoweredBy" value="1" <% if strShowImagePoweredBy = "1" then Response.Write("checked") %>> 
    文字：<input type="radio" class="radio" name="strShowImagePoweredBy" value="0" <% if strShowImagePoweredBy = "0" then Response.Write("checked") %>>
<!--    <a href="JavaScript:openWindow3('pop_config_help.asp#GfxButtons')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>-->
    </font></td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" colspan="2" align="center"><input type="submit" value="提交设定" id="submit1" name="submit1"> <input type="reset" value="重新设置" id="reset1" name="reset1"></td>
  </tr>
</table>
</form>
</font>
<%	end if %>
<!--#INCLUDE file="inc_footer.asp" -->
<% Else %>
<% Response.Redirect "admin_login.asp" %>
<% End IF %>
