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
    <img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="default.asp">返回论坛首页</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="admin_home.asp">论坛管理中心</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" height=15 width=15 border="0">&nbsp;NT功能设定<br>
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
			strSql = strSql & " SET "

			strSql = strSql & " C_STRNTGROUPS = " & Request.Form("strNTGroups")
			strSql = strSql & ", C_STRAUTOLOGON = " & Request.Form("strAutoLogon")
			strSql = strSql & " WHERE CONFIG_ID = " & 1

			my_Conn.Execute (strSql)
			Application(strCookieURL & "ConfigLoaded") = ""

%>
<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">设定已变更</font></p>
<meta http-equiv="Refresh" content="2; URL=admin_home.asp">

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">修改完成！</font></p>

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="admin_home.asp">回到管理主选单</font></a></p>
<%		else %>
	<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">你输入的资料有问题或遗漏</font></p>

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
<form action="admin_config_NT_features.asp" method="post" id="Form1" name="Form1">
<input type="hidden" name="Method_Type" value="Write_Configuration">
<table border="0" cellspacing="1" cellpadding="4" align=center bgcolor="<% =strPopUpBorderColor %>">
  <tr valign="top">
    <td bgcolor="<% =strHeadCellColor %>" colspan="2"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">NT功能设定</font></td>
  </tr>
<% if strAuthType = "nt" then %>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>Use NT Groups:</b>&nbsp;</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    开启：<input type="radio" class="radio" name="strNTGroups" value="1" <% if strNTGroups = "1" then Response.Write("checked") %>> 
    关闭：<input type="radio" class="radio" name="strNTGroups" value="0" <% if strNTGroups = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </font></td>
  </tr>
<% end if %>
<% if strAuthType = "nt" then %>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>Use NT AutoLogon:</b>&nbsp;</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    开启：<input type="radio" class="radio" name="strAutoLogon" value="1" <% if strAutoLogon = "1" then Response.Write("checked") %>> 
    关闭：<input type="radio" class="radio" name="strAutoLogon" value="0" <% if strAutoLogon = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </font></td>
  </tr>
<% end if %>

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
