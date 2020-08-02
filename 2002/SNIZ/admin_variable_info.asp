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
<!--#INCLUDE FILE="inc_functions.asp" -->
<!--#INCLUDE file="inc_top.asp" -->

	<table border="0" align=center width="100%">
	<tr>
	    <td width="33%" align="left" nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
		<img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="default.asp">返回论坛首页</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="admin_home.asp">论坛管理中心</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" height=15 width=15 border="0">&nbsp;论坛系统资讯<br>
		</font></td>
	</tr>
	</table><br>
    <table border="0" cellspacing="1" cellpadding="4" align=center width="100%" bgcolor="<% =strPopUpBorderColor %>">
     <tr>
        <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>Variable Name</b></font></td>
        <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>Value</b></font></td>
      </tr>

     <tr>
        <td align="center" colspan="2" bgcolor="<% =strHeadCellColor %>"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">General information</font></b></td>
     </tr>	
     <tr>
        <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>StrCookieUrl</b></font></td>
        <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><%=ChkString(StrCookieUrl, "display")%></font></td>
      </tr>
     <tr>
        <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>StrUniqueID</b></font></td>
        <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><%=ChkString(StrUniqueID, "display")%></font></td>
      </tr>
     <tr>
        <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>strAuthType</b></font></td>
        <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><%=ChkString(strAuthType, "display")%></font></td>
      </tr>
     <tr>
        <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>strDBNTSQLName</b></font></td>
        <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><%=ChkString(strDBNTSQLName, "display")%></font></td>
      </tr>
     <tr>
        <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>STRdbntUserName</b></font></td>
        <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><%=ChkString(STRdbntUserName, "display")%></font></td>
      </tr>

     <tr>
        <td align="center" colspan="2" bgcolor="<% =strHeadCellColor %>"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">Cookies</font></b></td>
      </tr>
<% for each key in Request.Cookies 

	if Request.Cookies(key).HasKeys then
		for each subkey in Request.Cookies(key)
%>
 		     <tr>
		        <td bgColor="<% =strPopUpTableColor %>" valign="top"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b><% =ChkString(key, "display") %> (<% =ChkString(subkey, "display") %>)</b></font></td>
		        <td bgColor="<% =strPopUpTableColor %>"><font face="courier" size="<% =strDefaultFontSize %>">
<%
		if Request.Cookies(key)(subkey) = "" then
			Response.Write "&nbsp;"
		else
			Response.Write ChkString(CStr(Request.Cookies(key)(subkey)), "display")
		end if 
%>
		        </font></td>
		      </tr>
<%		next
	else
%>
 		     <tr>
		        <td bgColor="<% =strPopUpTableColor %>" valign="top"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b><% =ChkString(key, "display") %></b></font></td>
		        <td bgColor="<% =strPopUpTableColor %>"><font face="courier" size="<% =strDefaultFontSize %>">
<%
		if Request.Cookies(key) = "" then
			Response.Write "&nbsp;"
		else
			Response.Write ChkString(CStr(Request.Cookies(key)), "display")
		end if 
%>
		        </font></td>
		      </tr>
<%
	end if
next  %>
 
     <tr>
        <td align="center" colspan="2" bgcolor="<% =strHeadCellColor %>"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">Session variables</font></b></td>
      </tr>
<% for each key in Session.Contents

	if left(key, len(strCookieUrl)) = strCookieUrl or left(key, len(strUniqueID)) = strUniqueID then
%>
      <tr>
        <td bgColor="<% =strPopUpTableColor %>" valign="top"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b><% =ChkString(key, "display") %></b></font></td>
        <td bgColor="<% =strPopUpTableColor %>"><font face="courier" size="<% =strDefaultFontSize %>">
<%
	if Session.Contents(key) = "" then
		Response.Write "&nbsp;"
	else
		Response.Write ChkString(CStr(Session.Contents(key)), "display")
	end if 
%>
        </font></td>
      </tr>
<% 
	end if
next 

%>

      <tr>
        <td align="center" colspan="2" bgcolor="<% =strHeadCellColor %>"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">Application variables</font></b></td>
      </tr>
<% for each key in Application.Contents

	if left(key, len(strCookieUrl)) = strCookieUrl or left(key, len(strUniqueID)) = strUniqueID then
%>
      <tr>
        <td bgColor="<% =strPopUpTableColor %>" valign="top"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b><% = ChkString(key, "display") %></b></font></td>
        <td bgColor="<% =strPopUpTableColor %>"><font face="courier" size="<% =strDefaultFontSize %>">
<%
	if Application.Contents(key) = "" then
		Response.Write "&nbsp;"
	else
		Response.Write ChkString(CStr(Application.Contents(key)), "display")
	end if 
%>
        </font></td>
      </tr>
<% 
	end if
next 

%>
    </table>
</p>
<!--#INCLUDE file="inc_footer.asp" -->
<% Else %>
<% Response.Redirect "admin_login.asp" %>
<% End IF %>
