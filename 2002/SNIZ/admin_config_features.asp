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
    <img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="default.asp">返回论坛首页</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="admin_home.asp">论坛管理中心</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" height=15 width=15 border="0">&nbsp;论坛设置<br>
    </font></td>
  </tr>
</table>
<%
	if Request.Form("Method_Type") = "Write_Configuration" then 
		Err_Msg = ""
		if Request.Form("strIMGInPosts") = "1" and Request.Form("strAllowForumCode") = "0" then 
			Err_Msg = Err_Msg & "<li>如果允许贴图就必须开启本版专用代码</li>"
		end if
		if (Request.Form("strHotTopic") = "1" and strHotTopic = "1") or (Request.Form("strHotTopic") = "1" and strHotTopic = "0") then
			if Request.Form("intHotTopicNum") = "" then 
				Err_Msg = Err_Msg & "<li>你必须设定热门主题的文章数</li>"
			end if
			if left(Request.Form("intHotTopicNum"), 1) = "-" then 
				Err_Msg = Err_Msg & "<li>你必须设定确切的热门主题文章数</li>"
			end if
			if left(Request.Form("intHotTopicNum"), 1) = "+" then 
				Err_Msg = Err_Msg & "<li>只能输入的数字不要包含 <b>+</b></li>"
			end if
		end if

		if Err_Msg = "" then

			'## Forum_SQL
			strSql = "UPDATE " & strTablePrefix & "CONFIG "
			strSql = strSql & " SET C_STRMOVETOPICMODE         = " & Request.Form("strMoveTopicMode") & ""
			strSql = strSql & ",    C_STRIPLOGGING             = " & Request.Form("strIPLogging") & ""
			strSql = strSql & ",    C_STRPRIVATEFORUMS         = " & Request.Form("strPrivateForums") & ""
			strSql = strSql & ",    C_STRSHOWMODERATORS        = " & Request.Form("strShowModerators") & ""
			strSql = strSql & ",    C_STRALLOWFORUMCODE        = " & Request.Form("strAllowForumCode") & ""
			strSql = strSql & ",    C_STRIMGINPOSTS            = " & Request.Form("strIMGInPosts") & ""
			strSql = strSql & ",    C_STRALLOWHTML             = " & Request.Form("strAllowHTML") & ""
			strSql = strSql & ",    C_STRHOTTOPIC              = " & Request.Form("strHotTopic") & ""
			if (Request.Form("strHotTopic") = "1" and strHotTopic = "1") or (Request.Form("strHotTopic") = "1" and strHotTopic = "0") then
				strSql = strSql & ",    C_INTHOTTOPICNUM           = " & Request.Form("intHotTopicNum") & ""
			end if
			strSql = strSql & ",    C_STRiconS                 = " & Request.Form("stricons") & ""
			strSql = strSql & ",    C_STRSECUREADMIN           = " & Request.Form("strSecureAdmin") & ""
			strSql = strSql & ",    C_STRNOCOOKIES             = " & Request.Form("strNoCookies") & ""
			strSql = strSql & ",    C_STREDITEDBYDATE          = " & Request.Form("strEditedByDate") & ""
			strSql = strSql & ",    C_STRSHOWSTATISTICS        = " & Request.Form("strShowStatistics") & ""
			strSql = strSql & ",    C_STRSHOWPAGING            = " & Request.Form("strShowPaging") & ""
			strSql = strSql & ",    C_STRSHOWTOPICNAV          = " & Request.Form("strShowTopicNav") & ""
			strSql = strSql & ",    C_STRPAGESIZE			   = " & Request.Form("strPageSize") & ""
			strSql = strSql & ",    C_STRPAGENUMBERSIZE        = " & Request.Form("strPageNumberSize") & ""
			strSql = strSql & " WHERE CONFIG_ID = " & 1

			my_Conn.Execute (strSql)
			Application(strCookieURL & "ConfigLoaded") = ""

%>
<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">论坛设置完毕</font></p>
<meta http-equiv="Refresh" content="0; URL=admin_home.asp">

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">修改完成！</font></p>

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="admin_home.asp">回到管理中心</font></a></p>
<%		else %>
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
<%		end if %>
<%	else %>
<form action="admin_config_features.asp" method="post" id="Form1" name="Form1">
<input type="hidden" name="Method_Type" value="Write_Configuration">
<table border="0" cellspacing="1" cellpadding="3" align=center bgcolor="<% =strPopUpBorderColor %>">
  <tr valign="top">
    <td bgcolor="<% =strHeadCellColor %>" colspan="2"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">论坛设置</font></td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">限制版主移动该版主题：</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    开启：<input type="radio" class="radio" name="strMoveTopicMode" value="1"<% if strMoveTopicMode <> "0" then Response.Write(" checked") %>> 
    关闭：<input type="radio" class="radio" name="strMoveTopicMode" value="0"<% if strMoveTopicMode = "0" then Response.Write(" checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#MoveTopicMode')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">记录IP：</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    开启：<input type="radio" class="radio" name="strIPLogging" value="1"<% if strIPLogging <> "0" then Response.Write(" checked") %>> 
    关闭：<input type="radio" class="radio" name="strIPLogging" value="0"<% if strIPLogging = "0" then Response.Write(" checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#IPLogging')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">秘密论坛：</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    开启：<input type="radio" class="radio" name="strPrivateForums" value="1"<% if strPrivateForums <> "0" then Response.Write(" checked") %>> 
    关闭：<input type="radio" class="radio" name="strPrivateForums" value="0"<% if strPrivateForums = "0" then Response.Write(" checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#privateforums')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">显示版主：</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    开启：<input type="radio" class="radio" name="strShowModerators" value="1"<% if strShowModerators <> "0" then Response.Write(" checked") %>> 
    关闭：<input type="radio" class="radio" name="strShowModerators" value="0"<% if strShowModerators = "0" then Response.Write(" checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#ShowModerator')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">本版专用代码：&nbsp;</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    开启：<input type="radio" class="radio" name="strAllowForumCode" value="1"<% if strAllowForumCode <> "0" then Response.Write(" checked") %>> 
    关闭：<input type="radio" class="radio" name="strAllowForumCode" value="0"<% if strAllowForumCode = "0" then Response.Write(" checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#AllowForumCode')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">贴图功能：</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    开启：<input type="radio" class="radio" name="strIMGInPosts" value="1" <% if (lcase(strIMGInPosts) <> "0") then Response.Write("checked")%>> 
    关闭：<input type="radio" class="radio" name="strIMGInPosts" value="0" <% if (lcase(strIMGInPosts) = "0") then Response.Write("checked")%>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#imginposts')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
   </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">允许HTML代码：</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    开启：<input type="radio" class="radio" name="strAllowHTML" value="1"<% if strAllowHTML <> "0" then Response.Write(" checked") %>> 
    关闭：<input type="radio" class="radio" name="strAllowHTML" value="0"<% if strAllowHTML = "0" then Response.Write(" checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#AllowHTML')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">保全管理模式：</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    开启：<input type="radio" class="radio" name="strSecureAdmin" value="1" <% if lcase(strSecureAdmin) <> "0" then Response.Write("checked") %>> 
    关闭：<input type="radio" class="radio" name="strSecureAdmin" value="0" <% if lcase(strSecureAdmin) = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#secureadminmode')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">无 Cookie 模式：</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    开启：<input type="radio" class="radio" name="strNoCookies" value="1" <% if lcase(strNoCookies) <> "0" then Response.Write("checked") %>> 
    关闭：<input type="radio" class="radio" name="strNoCookies" value="0" <% if lcase(strNoCookies) = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#allownoncookies')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">显示重编日期：</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    开启：<input type="radio" class="radio" name="strEditedByDate" value="1" <% if lcase(strEditedByDate) <> "0" then Response.Write("checked") %>> 
    关闭：<input type="radio" class="radio" name="strEditedByDate" value="0" <% if lcase(strEditedByDate) = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#editedbydate')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">热门主题：</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    开启：<input type="radio" class="radio" name="strHotTopic" value="1" <% if (strHotTopic <> "0" or lcase(HotTopic) <> "0") then Response.Write("checked") %>> 
    关闭：<input type="radio" class="radio" name="strHotTopic" value="0" <% if (strHotTopic = "0" or lcase(HotTopic) = "0") then Response.Write("checked") %>>
    <input type="text" name="intHotTopicNum" size="5" value="<% if intHotTopicNum <> "" then Response.Write(intHotTopicNum) else Response.Write("20") %>">
    <a href="JavaScript:openWindow3('pop_config_help.asp#hottopics')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">表情转换图示：</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    开启：<input type="radio" class="radio" name="stricons" value="1" <% if (lcase(stricons) <> "0" or lcase(Smiles) <> "0") then Response.Write("checked") %>> 
    关闭：<input type="radio" class="radio" name="stricons" value="0" <% if (lcase(stricons) = "0" or lcase(Smiles) = "0") then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#<%=strImageURL %>icons')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
   <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">详细状态列：</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    开启：<input type="radio" class="radio" name="strShowStatistics" value="1" <% if strShowStatistics <> "0" then Response.Write("checked") %>> 
    关闭：<input type="radio" class="radio" name="strShowStatistics" value="0" <% if strShowStatistics = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#stats')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">页面快速连接：</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    开启：<input type="radio" class="radio" name="strShowPaging" value="1" <% if strShowPaging <> "0" then Response.Write("checked") %>> 
    关闭：<input type="radio" class="radio" name="strShowPaging" value="0" <% if strShowPaging = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#ShowPaging')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">显示 前/后 主题：</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    开启：<input type="radio" class="radio" name="strShowTopicNav" value="1" <% if strShowTopicNav <> "0" then Response.Write("checked") %>> 
    关闭：<input type="radio" class="radio" name="strShowTopicNav" value="0" <% if strShowTopicNav = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#ShowTopicNav')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">每页显示主题数：</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    <input type="text" name="strPageSize" size="5" value="<% if strPageSize <> "" then Response.Write(strPageSize) else Response.Write("15") %>">
<!--    <a href="JavaScript:openWindow3('pop_config_help.asp#hottopics')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>-->
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">每行显示主题数：</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    <input type="text" name="strPageNumberSize" size="5" value="<% if strPageNumberSize <> "" then Response.Write(strPageNumberSize) else Response.Write("10") %>">
<!--    <a href="JavaScript:openWindow3('pop_config_help.asp#hottopics')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>-->
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" colspan="2" align="center"><input type="submit" value="提交设置" id="submit1" name="submit1"> <input type="reset" value="重新设置" id="reset1" name="reset1"></td>
  </tr>
</table>
</form>
</font>
<%	end if %>
<!--#INCLUDE file="inc_footer.asp" -->
<% Else %>
<% Response.Redirect "admin_login.asp" %>
<% End IF %>
