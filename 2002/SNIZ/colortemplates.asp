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
<%
t_select=Request.Form("Select")
if request.Form("template")="Write" then
If t_select = "darkstation" then
Response.Write t_select & request.form("template")
strSql = "UPDATE " & strTablePrefix & "CONFIG "
strSql = strSql & " SET C_STRDEFAULTFONTFACE = '宋体, Arial, Helvetica', C_STRDEFAULTFONTSIZE='2', C_STRHEADERFONTSIZE = '3', "
strSql = strSql & " C_STRFOOTERFONTSIZE = '1', C_STRPAGEBGCOLOR = 'FFFFFF', C_STRDEFAULTFONTCOLOR = '000000', "
strSql = strSql & " C_STRLINKCOLOR = '000066', C_STRLINKTEXTDECORATION = 'none', C_STRVISITEDLINKCOLOR = '000066', "
strSql = strSql & " C_STRVISITEDTEXTDECORATION = 'none', C_STRACTIVELINKCOLOR = '990000', C_STRHOVERFONTCOLOR = '990000', "
strSql = strSql & " C_STRHOVERTEXTDECORATION = 'underline', C_STRHEADCELLCOLOR = 'orange', C_STRHEADFONTCOLOR = 'FFFFFF', "
strSql = strSql & " C_STRCATEGORYCELLCOLOR = 'orange', C_STRCATEGORYFONTCOLOR = 'FFFFFF', C_STRFORUMFIRSTCELLCOLOR = 'F7EFD6', "
strSql = strSql & " C_STRFORUMCELLCOLOR = 'F7EFD6', C_STRALTFORUMCELLCOLOR = 'F7FFE7', C_STRFORUMFONTCOLOR = '000000', "
strSql = strSql & " C_STRFORUMLINKCOLOR = '000066', C_STRTABLEBORDERCOLOR = '000000', C_STRPOPUPTABLECOLOR = 'F7FFE7', "
strSql = strSql & " C_STRPOPUPBORDERCOLOR = '000000', C_STRNEWFONTCOLOR = '990000', C_STRTOPICWIDTHLEFT = '100', "
strSql = strSql & " C_STRTOPICNOWRAPLEFT = '1', C_STRTOPICWIDTHRIGHT = '100%', C_STRTOPICNOWRAPRIGHT = '0' WHERE CONFIG_ID = " & 1
End If
If t_select = "beige" then
strSql = "UPDATE " & strTablePrefix & "CONFIG "
strSql = strSql & " SET C_STRDEFAULTFONTFACE = '宋体, Arial, Helvetica', C_STRDEFAULTFONTSIZE = '2', C_STRHEADERFONTSIZE = '4', "
strSql = strSql & " C_STRFOOTERFONTSIZE = '1', C_STRPAGEBGCOLOR = 'CCCCBB', C_STRDEFAULTFONTCOLOR = '000066', "
strSql = strSql & " C_STRLINKCOLOR = '000099', C_STRLINKTEXTDECORATION = 'underline', C_STRVISITEDLINKCOLOR = '000066', "
strSql = strSql & " C_STRVISITEDTEXTDECORATION = 'underline', C_STRACTIVELINKCOLOR = '990000', C_STRHOVERFONTCOLOR = '990000', "
strSql = strSql & " C_STRHOVERTEXTDECORATION = 'underline', C_STRHEADCELLCOLOR = '808060', C_STRHEADFONTCOLOR = 'F9F9F9', "
strSql = strSql & " C_STRCATEGORYCELLCOLOR = '808060', C_STRCATEGORYFONTCOLOR = 'EEEEEE', C_STRFORUMFIRSTCELLCOLOR = 'B9B9A4', "
strSql = strSql & " C_STRFORUMCELLCOLOR = 'B9B9A4', C_STRALTFORUMCELLCOLOR = 'CACAB9', C_STRFORUMFONTCOLOR = '000099', "
strSql = strSql & " C_STRFORUMLINKCOLOR = '000066', C_STRTABLEBORDERCOLOR = '000000', C_STRPOPUPTABLECOLOR = 'F9F9F9', "
strSql = strSql & " C_STRPOPUPBORDERCOLOR = '000000', C_STRNEWFONTCOLOR = '000066', C_STRTOPICWIDTHLEFT = '100', "
strSql = strSql & " C_STRTOPICNOWRAPLEFT = '1', C_STRTOPICWIDTHRIGHT = '100%', C_STRTOPICNOWRAPRIGHT = '0' "
strSql = strSql & " WHERE CONFIG_ID = " & 1
End If
If t_select = "redrose" then
strSql = "UPDATE " & strTablePrefix & "CONFIG "
strSql = strSql & " SET C_STRDEFAULTFONTFACE = '宋体, Arial, Helvetica', C_STRDEFAULTFONTSIZE = '2', C_STRHEADERFONTSIZE = '4', "
strSql = strSql & " C_STRFOOTERFONTSIZE = '1', C_STRPAGEBGCOLOR = 'firebrick', C_STRDEFAULTFONTCOLOR = '000000', "
strSql = strSql & " C_STRLINKCOLOR  = '800000', C_STRLINKTEXTDECORATION = 'blink', C_STRVISITEDLINKCOLOR = 'maroon', "
strSql = strSql & " C_STRVISITEDTEXTDECORATION = 'underline', C_STRACTIVELINKCOLOR = 'lightsalmon', C_STRHOVERFONTCOLOR = 'mistyrose', "
strSql = strSql & " C_STRHOVERTEXTDECORATION = 'blink', C_STRHEADCELLCOLOR = '800000', C_STRHEADFONTCOLOR = 'antiquewhite', "
strSql = strSql & " C_STRCATEGORYCELLCOLOR = 'tomato', C_STRCATEGORYFONTCOLOR = '000000', C_STRFORUMFIRSTCELLCOLOR = 'lightcoral', "
strSql = strSql & " C_STRFORUMCELLCOLOR = 'lightcoral', C_STRALTFORUMCELLCOLOR = 'indianred', C_STRFORUMFONTCOLOR = '000000', "
strSql = strSql & " C_STRFORUMLINKCOLOR = '800000', C_STRTABLEBORDERCOLOR = '000000', C_STRPOPUPTABLECOLOR = 'F9F9F9', "
strSql = strSql & " C_STRPOPUPBORDERCOLOR = '000000', C_STRNEWFONTCOLOR = '000000', C_STRTOPICWIDTHLEFT = '100', "
strSql = strSql & " C_STRTOPICNOWRAPLEFT = '1', C_STRTOPICWIDTHRIGHT = '100%', C_STRTOPICNOWRAPRIGHT = '0' "
strSql = strSql & " WHERE CONFIG_ID = " & 1
End If
If t_select = "grey" then
strSql = "UPDATE " & strTablePrefix & "CONFIG "
strSql = strSql & " SET C_STRDEFAULTFONTFACE = '宋体, Arial, Helvetica', C_STRDEFAULTFONTSIZE = '2', "
strSql = strSql & " C_STRHEADERFONTSIZE = '4', C_STRFOOTERFONTSIZE = '1', C_STRPAGEBGCOLOR = '454542', "
strSql = strSql & " C_STRDEFAULTFONTCOLOR = 'A9A9A9', C_STRLINKCOLOR = '858693', C_STRLINKTEXTDECORATION = 'underline', "
strSql = strSql & " C_STRVISITEDLINKCOLOR = '858693', C_STRVISITEDTEXTDECORATION = 'underline', C_STRACTIVELINKCOLOR = 'C0C0C0', "
strSql = strSql & " C_STRHOVERFONTCOLOR = 'C0C0C0', C_STRHOVERTEXTDECORATION = 'underline', C_STRHEADCELLCOLOR = '616160', "
strSql = strSql & " C_STRHEADFONTCOLOR = 'A9A9A9', C_STRCATEGORYCELLCOLOR = '696969', C_STRCATEGORYFONTCOLOR = 'A9A9A9', "
strSql = strSql & " C_STRFORUMFIRSTCELLCOLOR = '474949', C_STRFORUMCELLCOLOR = '474949', C_STRALTFORUMCELLCOLOR = '404040', "
strSql = strSql & " C_STRFORUMFONTCOLOR = 'A9A9A9', C_STRFORUMLINKCOLOR = '000000', C_STRTABLEBORDERCOLOR = 'A9A9A9', "
strSql = strSql & " C_STRPOPUPTABLECOLOR = '808080', C_STRPOPUPBORDERCOLOR = '808080', C_STRNEWFONTCOLOR = '000066', "
strSql = strSql & " C_STRTOPICWIDTHLEFT = '100', C_STRTOPICNOWRAPLEFT = '1', C_STRTOPICWIDTHRIGHT = '100%', C_STRTOPICNOWRAPRIGHT = '0' "
strSql = strSql & " WHERE CONFIG_ID = " & 1
End If
If t_select = "BPISG" then
strSql = "UPDATE " & strTablePrefix & "CONFIG "
strSql = strSql & " SET C_STRDEFAULTFONTFACE = '宋体, Arial, Helvetica', C_STRDEFAULTFONTSIZE = '2', C_STRHEADERFONTSIZE = '4', "
strSql = strSql & " C_STRFOOTERFONTSIZE = '1', C_STRPAGEBGCOLOR = '000066', C_STRDEFAULTFONTCOLOR = 'FFFFFF', C_STRLINKCOLOR = '9898D0', "
strSql = strSql & " C_STRLINKTEXTDECORATION = 'underline', C_STRVISITEDLINKCOLOR = 'GRAY', C_STRVISITEDTEXTDECORATION = 'underline', "
strSql = strSql & " C_STRACTIVELINKCOLOR = '9898D0', C_STRHOVERFONTCOLOR = '9898D0', C_STRHOVERTEXTDECORATION = 'underline', "
strSql = strSql & " C_STRHEADCELLCOLOR = '000066', C_STRHEADFONTCOLOR = 'F9F9F9', C_STRCATEGORYCELLCOLOR = '000066', "
strSql = strSql & " C_STRCATEGORYFONTCOLOR = 'FFFFFF', C_STRFORUMFIRSTCELLCOLOR = '000066', C_STRFORUMCELLCOLOR = '000066', "
strSql = strSql & " C_STRALTFORUMCELLCOLOR = '000066', C_STRFORUMFONTCOLOR = 'FFFFFF', C_STRFORUMLINKCOLOR = '9898D0', "
strSql = strSql & " C_STRTABLEBORDERCOLOR = 'FFFFFF', C_STRPOPUPTABLECOLOR = '000066', C_STRPOPUPBORDERCOLOR = 'FFFFFF', "
strSql = strSql & " C_STRNEWFONTCOLOR = 'FFFFFF', C_STRTOPICWIDTHLEFT = '100', C_STRTOPICNOWRAPLEFT = '1', C_STRTOPICWIDTHRIGHT = '100%', "
strSql = strSql & " C_STRTOPICNOWRAPRIGHT = '0' WHERE CONFIG_ID = " & 1
End If
If t_select = "green" then
strSql = "UPDATE " & strTablePrefix & "CONFIG "
strSql = strSql & " SET C_STRDEFAULTFONTFACE = '宋体, Arial, Helvetica', C_STRDEFAULTFONTSIZE = '2', C_STRHEADERFONTSIZE = '4', "
strSql = strSql & " C_STRFOOTERFONTSIZE = '1', C_STRPAGEBGCOLOR = '003333', C_STRDEFAULTFONTCOLOR = 'A9A9A9', C_STRLINKCOLOR = '009999', "
strSql = strSql & " C_STRLINKTEXTDECORATION = 'underline', C_STRVISITEDLINKCOLOR = '339999', C_STRVISITEDTEXTDECORATION = 'underline', "
strSql = strSql & " C_STRACTIVELINKCOLOR = 'FF0000', C_STRHOVERFONTCOLOR = 'FF0000', C_STRHOVERTEXTDECORATION = 'underline', "
strSql = strSql & " C_STRHEADCELLCOLOR = '003333', C_STRHEADFONTCOLOR = 'F9F9F9', C_STRCATEGORYCELLCOLOR = '006666', "
strSql = strSql & " C_STRCATEGORYFONTCOLOR = 'EEEEEE', C_STRFORUMFIRSTCELLCOLOR = 'EFEFEF', C_STRFORUMCELLCOLOR = 'EFEFEF', "
strSql = strSql & " C_STRALTFORUMCELLCOLOR = 'DFDFDF', C_STRFORUMFONTCOLOR = '000000', C_STRFORUMLINKCOLOR = '0000CC', "
strSql = strSql & " C_STRTABLEBORDERCOLOR = '000000', C_STRPOPUPTABLECOLOR = 'F9F9F9', C_STRPOPUPBORDERCOLOR = '000000', "
strSql = strSql & " C_STRNEWFONTCOLOR = '000066', C_STRTOPICWIDTHLEFT = '100', C_STRTOPICNOWRAPLEFT = '1', "
strSql = strSql & " C_STRTOPICWIDTHRIGHT = '100%', C_STRTOPICNOWRAPRIGHT = '0' WHERE CONFIG_ID = " & 1
End If
If t_select = "default" then
strSql = "UPDATE " & strTablePrefix & "CONFIG "
strSql = strSql & " SET C_STRDEFAULTFONTFACE = '宋体, Arial, Helvetica', C_STRDEFAULTFONTSIZE = '2', C_STRHEADERFONTSIZE = '4', "
strSql = strSql & " C_STRFOOTERFONTSIZE = '1', C_STRPAGEBGCOLOR = 'FFFFFF', C_STRDEFAULTFONTCOLOR = 'midnightblue', C_STRLINKCOLOR = 'darkblue', "
strSql = strSql & " C_STRLINKTEXTDECORATION = 'underline', C_STRVISITEDLINKCOLOR = 'blue', C_STRVISITEDTEXTDECORATION = 'underline', "
strSql = strSql & " C_STRACTIVELINKCOLOR = 'red', C_STRHOVERFONTCOLOR = 'red', C_STRHOVERTEXTDECORATION = 'underline', "
strSql = strSql & " C_STRHEADCELLCOLOR = 'midnightblue', C_STRHEADFONTCOLOR = 'mintcream', C_STRCATEGORYCELLCOLOR = 'slateblue', "
strSql = strSql & " C_STRCATEGORYFONTCOLOR = 'mintcream', C_STRFORUMFIRSTCELLCOLOR = 'whitesmoke', C_STRFORUMCELLCOLOR = 'whitesmoke', "
strSql = strSql & " C_STRALTFORUMCELLCOLOR = 'gainsboro', C_STRFORUMFONTCOLOR = 'midnightblue', C_STRFORUMLINKCOLOR = 'darkblue', "
strSql = strSql & " C_STRTABLEBORDERCOLOR = '000000', C_STRPOPUPTABLECOLOR = 'navyblue', C_STRPOPUPBORDERCOLOR = '000000', "
strSql = strSql & " C_STRNEWFONTCOLOR = 'blue', C_STRTOPICWIDTHLEFT = '100', C_STRTOPICNOWRAPLEFT = '1', "
strSql = strSql & " C_STRTOPICWIDTHRIGHT = '100%', C_STRTOPICNOWRAPRIGHT = '0' WHERE CONFIG_ID = " & 1
End If
If t_select = "slide" then
strSql = "UPDATE " & strTablePrefix & "CONFIG "
strSql = strSql & " SET C_STRDEFAULTFONTFACE = '宋体, Arial, Helvetica', C_STRDEFAULTFONTSIZE = '2', C_STRHEADERFONTSIZE = '4', "
strSql = strSql & " C_STRFOOTERFONTSIZE = '1', C_STRPAGEBGCOLOR = 'slategrey', C_STRDEFAULTFONTCOLOR = 'lightsteelblue', C_STRLINKCOLOR = 'azure', "
strSql = strSql & " C_STRLINKTEXTDECORATION = 'underline', C_STRVISITEDLINKCOLOR = 'orange', C_STRVISITEDTEXTDECORATION = 'underline', "
strSql = strSql & " C_STRACTIVELINKCOLOR = 'yellowgreen', C_STRHOVERFONTCOLOR = 'red', C_STRHOVERTEXTDECORATION = 'underline', "
strSql = strSql & " C_STRHEADCELLCOLOR = 'dimgray', C_STRHEADFONTCOLOR = 'black', C_STRCATEGORYCELLCOLOR = 'red', "
strSql = strSql & " C_STRCATEGORYFONTCOLOR = 'black', C_STRFORUMFIRSTCELLCOLOR = 'red', C_STRFORUMCELLCOLOR = 'black', "
strSql = strSql & " C_STRALTFORUMCELLCOLOR = 'blue', C_STRFORUMFONTCOLOR = 'azure', C_STRFORUMLINKCOLOR = 'lightgreen', "
strSql = strSql & " C_STRTABLEBORDERCOLOR = 'darkgray', C_STRPOPUPTABLECOLOR = 'red', C_STRPOPUPBORDERCOLOR = '000000', "
strSql = strSql & " C_STRNEWFONTCOLOR = 'red', C_STRTOPICWIDTHLEFT = '100', C_STRTOPICNOWRAPLEFT = '1', "
strSql = strSql & " C_STRTOPICWIDTHRIGHT = '100%', C_STRTOPICNOWRAPRIGHT = '0' WHERE CONFIG_ID = " & 1
End If
If t_select = "fantasy" then
strSql = "UPDATE " & strTablePrefix & "CONFIG "
strSql = strSql & " SET C_STRDEFAULTFONTFACE = '宋体, Arial, Helvetica', C_STRDEFAULTFONTSIZE = '2', C_STRHEADERFONTSIZE = '4', "
strSql = strSql & " C_STRFOOTERFONTSIZE = '1', C_STRPAGEBGCOLOR = '000000', C_STRDEFAULTFONTCOLOR = 'mintcream', C_STRLINKCOLOR = 'red', "
strSql = strSql & " C_STRLINKTEXTDECORATION = 'underline', C_STRVISITEDLINKCOLOR = 'red', C_STRVISITEDTEXTDECORATION = 'underline', "
strSql = strSql & " C_STRACTIVELINKCOLOR = 'red', C_STRHOVERFONTCOLOR = 'firebrick', C_STRHOVERTEXTDECORATION = 'underline', "
strSql = strSql & " C_STRHEADCELLCOLOR = 'lightslateblue', C_STRHEADFONTCOLOR = 'mintcream', C_STRCATEGORYCELLCOLOR = 'midnightblue', "
strSql = strSql & " C_STRCATEGORYFONTCOLOR = 'mintcream', C_STRFORUMFIRSTCELLCOLOR = 'lightgray', C_STRFORUMCELLCOLOR = 'midnightblue', "
strSql = strSql & " C_STRALTFORUMCELLCOLOR = 'darkblue', C_STRFORUMFONTCOLOR = 'mintcream', C_STRFORUMLINKCOLOR = 'red', "
strSql = strSql & " C_STRTABLEBORDERCOLOR = '000000', C_STRPOPUPTABLECOLOR = 'midnightblue', C_STRPOPUPBORDERCOLOR = '000000', "
strSql = strSql & " C_STRNEWFONTCOLOR = 'red', C_STRTOPICWIDTHLEFT = '100', C_STRTOPICNOWRAPLEFT = '1', "
strSql = strSql & " C_STRTOPICWIDTHRIGHT = '100%', C_STRTOPICNOWRAPRIGHT = '0' WHERE CONFIG_ID = " & 1
End If
If t_select = "px" then
strSql = "UPDATE " & strTablePrefix & "CONFIG "
strSql = strSql & " SET C_STRDEFAULTFONTFACE = '宋体, Arial, Helvetica', C_STRDEFAULTFONTSIZE = '2', C_STRHEADERFONTSIZE = '4', "
strSql = strSql & " C_STRFOOTERFONTSIZE = '1', C_STRPAGEBGCOLOR = '263C6D', C_STRDEFAULTFONTCOLOR = '3399CC', C_STRLINKCOLOR = 'FFFFFF', "
strSql = strSql & " C_STRLINKTEXTDECORATION = 'none', C_STRVISITEDLINKCOLOR = '55CC55', C_STRVISITEDTEXTDECORATION = 'none', "
strSql = strSql & " C_STRACTIVELINKCOLOR = '990000', C_STRHOVERFONTCOLOR = 'DD0000', C_STRHOVERTEXTDECORATION = 'underline', "
strSql = strSql & " C_STRHEADCELLCOLOR = '1177AA', C_STRHEADFONTCOLOR = 'F9F9F9', C_STRCATEGORYCELLCOLOR = '006699', "
strSql = strSql & " C_STRCATEGORYFONTCOLOR = 'EEEEEE', C_STRFORUMFIRSTCELLCOLOR = '263C6D', C_STRFORUMCELLCOLOR = '263C6D', "
strSql = strSql & " C_STRALTFORUMCELLCOLOR = '263C6D', C_STRFORUMFONTCOLOR = '3399CC', C_STRFORUMLINKCOLOR = '000066', "
strSql = strSql & " C_STRTABLEBORDERCOLOR = '3399CC', C_STRPOPUPTABLECOLOR = '263C6D', C_STRPOPUPBORDERCOLOR = '3399CC', "
strSql = strSql & " C_STRNEWFONTCOLOR = '3399CC', C_STRTOPICWIDTHLEFT = '100', C_STRTOPICNOWRAPLEFT = '1', "
strSql = strSql & " C_STRTOPICWIDTHRIGHT = '100%', C_STRTOPICNOWRAPRIGHT = '0' WHERE CONFIG_ID = " & 1
End If
If t_select = "madonna" then
strSql = "UPDATE " & strTablePrefix & "CONFIG "
strSql = strSql & " SET C_STRDEFAULTFONTFACE = '宋体, Arial, Helvetica', C_STRDEFAULTFONTSIZE = '2', C_STRHEADERFONTSIZE = '4', "
strSql = strSql & " C_STRFOOTERFONTSIZE = '1', C_STRPAGEBGCOLOR = 'FFFFFF', C_STRDEFAULTFONTCOLOR = '000000', C_STRLINKCOLOR = '40A0A0', "
strSql = strSql & " C_STRLINKTEXTDECORATION = 'underline', C_STRVISITEDLINKCOLOR = '40A0A0', C_STRVISITEDTEXTDECORATION = 'underline', "
strSql = strSql & " C_STRACTIVELINKCOLOR = 'FF8080', C_STRHOVERFONTCOLOR = 'FF8080', C_STRHOVERTEXTDECORATION = 'underline', "
strSql = strSql & " C_STRHEADCELLCOLOR = '40A0A0', C_STRHEADFONTCOLOR = 'FFFFFF', C_STRCATEGORYCELLCOLOR = 'FFFFFF', "
strSql = strSql & " C_STRCATEGORYFONTCOLOR = '000000', C_STRFORUMFIRSTCELLCOLOR = 'F9F9F9', C_STRFORUMCELLCOLOR = 'F9F9F9', "
strSql = strSql & " C_STRALTFORUMCELLCOLOR = 'gainsboro', C_STRFORUMFONTCOLOR = '000000', C_STRFORUMLINKCOLOR = '40A0A0', "
strSql = strSql & " C_STRTABLEBORDERCOLOR = '000000', C_STRPOPUPTABLECOLOR = 'F9F9F9', C_STRPOPUPBORDERCOLOR = '000000', "
strSql = strSql & " C_STRNEWFONTCOLOR = '000000', C_STRTOPICWIDTHLEFT = '100', C_STRTOPICNOWRAPLEFT = '1', "
strSql = strSql & " C_STRTOPICWIDTHRIGHT = '100%', C_STRTOPICNOWRAPRIGHT = '0' WHERE CONFIG_ID = " & 1
End If
If t_select = "blacksun" then
strSql = "UPDATE " & strTablePrefix & "CONFIG "
strSql = strSql & " SET C_STRDEFAULTFONTFACE = '宋体, Arial, Helvetica', C_STRDEFAULTFONTSIZE = '2', C_STRHEADERFONTSIZE = '3', "
strSql = strSql & " C_STRFOOTERFONTSIZE = '1', C_STRPAGEBGCOLOR = '000000', C_STRDEFAULTFONTCOLOR = 'FFFFFF', C_STRLINKCOLOR = 'FFFFFF', "
strSql = strSql & " C_STRLINKTEXTDECORATION = 'underline', C_STRVISITEDLINKCOLOR = 'FFFFFF', C_STRVISITEDTEXTDECORATION = 'underline', "
strSql = strSql & " C_STRACTIVELINKCOLOR = 'FFFFFF', C_STRHOVERFONTCOLOR = 'FFFFFF', C_STRHOVERTEXTDECORATION = 'underline', "
strSql = strSql & " C_STRHEADCELLCOLOR = '27467A', C_STRHEADFONTCOLOR = 'FFFFFF', C_STRCATEGORYCELLCOLOR = 'D46B2A', "
strSql = strSql & " C_STRCATEGORYFONTCOLOR = 'FFFFFF', C_STRFORUMFIRSTCELLCOLOR = '264C8B', C_STRFORUMCELLCOLOR = '264C8B', "
strSql = strSql & " C_STRALTFORUMCELLCOLOR = '264C8B', C_STRFORUMFONTCOLOR = 'F9F9F9', C_STRFORUMLINKCOLOR = 'F9F9F9', "
strSql = strSql & " C_STRTABLEBORDERCOLOR = '000000', C_STRPOPUPTABLECOLOR = '264C8B', C_STRPOPUPBORDERCOLOR = '000000', "
strSql = strSql & " C_STRNEWFONTCOLOR = '264C8B', C_STRTOPICWIDTHLEFT = '100', C_STRTOPICNOWRAPLEFT = '1', "
strSql = strSql & " C_STRTOPICWIDTHRIGHT = '100%', C_STRTOPICNOWRAPRIGHT = '0' WHERE CONFIG_ID = " & 1
End If
my_Conn.Execute (strSql)
Application(strCookieURL & "ConfigLoaded") = ""
End if%>
<!--#INCLUDE file="inc_footer.asp" -->
<% Else %>
<% Response.Redirect "admin_login.asp" %> 
<% End IF 
Response.Redirect ("admin_config_colors.asp")%>
