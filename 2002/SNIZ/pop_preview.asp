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
<%
strMessagePreview = Request.Cookies("strMessagePreview")
Response.Cookies("strMessagePreview") = ""
if strMessagePreview = "" or IsNull(strMessagePreview) then
	strMessagePreview = "[center][b]<没有内容可以预览>[/b][/center]"
end if
%>
<!--#INCLUDE FILE="inc_top_short.asp"-->
<script language="JavaScript">
<!--
function jumpTo(s) {if (s.selectedIndex != 0) top.location.href = s.options[s.selectedIndex].value;return 1;}
// -->
</script>
<table border="0" width="95%" height="80%" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td bgcolor="<% =strTableBorderColor %>">
    <table border="0" width="100%" height="100%" cellspacing="1" cellpadding="4">
      <tr>
        <td align="center" bgcolor="<% =strHeadCellColor %>" width="<% =strTopicWidthRight %>" height="20" <% if lcase(strTopicNoWrapRight) = "1" then Response.Write(" nowrap") %>><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">预览</font></b></td>
	  </tr>
      <tr>
        <td bgcolor="<% =strForumCellColor %>" valign="top">
        <font color="<% =strForumFontColor %>" face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% =formatStr(chkString(strMessagePreview,"preview")) %></font>
        </td>
      </tr>
    </table>
    </td>
  </tr>
</table>
</div>
<!--#INCLUDE FILE="inc_footer_short.asp" -->