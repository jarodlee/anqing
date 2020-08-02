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
    <img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="default.asp">返回论坛首页</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="admin_home.asp">论坛管理中心</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" height=15 width=15 border="0">&nbsp;服务器日期/时间设定<br>
    </font></td>
  </tr>
</table>
<%
	if Request.Form("Method_Type") = "Write_Configuration" then 
		Err_Msg = ""

		if Err_Msg = "" then

			'## Forum_SQL
			strSql = "UPDATE " & strTablePrefix & "CONFIG "
			strSql = strSql & " SET C_STRDATETYPE              = '" & Request.Form("strDateType") & "', "
			strSql = strSql & "     C_STRTIMETYPE              = '" & Request.Form("strTimeType") & "', "
			strSql = strSql & "     C_STRTIMEADJUST            = " & Request.Form("strTimeAdjust") & " "
			strSql = strSql & " WHERE CONFIG_ID = " & 1
		
			my_Conn.Execute (strSql)
			Application(strCookieURL & "ConfigLoaded") = ""

%>
<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">服务器日期/时间设定已经更新完毕！</font></p>
<meta http-equiv="Refresh" content="2; URL=admin_home.asp">

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
<form action="admin_config_datetime.asp" method="post" id="Form1" name="Form1">
<input type="hidden" name="Method_Type" value="Write_Configuration">
<table border="0" cellspacing="1" cellpadding="4" align=center bgcolor="<% =strPopUpBorderColor %>">
  <tr valign="top">
    <td bgcolor="<% =strHeadCellColor %>" colspan="2"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">服务器日期/时间设定</font></td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">时间显示方式：</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    全日制(24) <input type="radio" class="radio" name="strTimeType" value="24" <% if strTimeType = "24" then Response.Write("checked") %>> 
    半日制(12) <input type="radio" class="radio" name="strTimeType" value="12" <% if strTimeType = "12" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#timetype')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </font></td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">时差：</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    <select name="strTimeAdjust">
      <option Value="-24"<% if (lcase(strTimeAdjust)="-24") then Response.Write(" selected") %>>-24</option>
      <option Value="-23"<% if (lcase(strTimeAdjust)="-23") then Response.Write(" selected") %>>-23</option>
      <option Value="-22"<% if (lcase(strTimeAdjust)="-22") then Response.Write(" selected") %>>-22</option>
      <option Value="-21"<% if (lcase(strTimeAdjust)="-21") then Response.Write(" selected") %>>-21</option>
      <option Value="-20"<% if (lcase(strTimeAdjust)="-20") then Response.Write(" selected") %>>-20</option>
      <option Value="-19"<% if (lcase(strTimeAdjust)="-19") then Response.Write(" selected") %>>-19</option>
      <option Value="-18"<% if (lcase(strTimeAdjust)="-18") then Response.Write(" selected") %>>-18</option>
      <option Value="-17"<% if (lcase(strTimeAdjust)="-17") then Response.Write(" selected") %>>-17</option>
      <option Value="-16"<% if (lcase(strTimeAdjust)="-16") then Response.Write(" selected") %>>-16</option>
      <option Value="-15"<% if (lcase(strTimeAdjust)="-15") then Response.Write(" selected") %>>-15</option>
      <option Value="-14"<% if (lcase(strTimeAdjust)="-14") then Response.Write(" selected") %>>-14</option>
      <option Value="-13"<% if (lcase(strTimeAdjust)="-13") then Response.Write(" selected") %>>-13</option>
      <option Value="-12"<% if (lcase(strTimeAdjust)="-12") then Response.Write(" selected") %>>-12</option>
      <option Value="-11"<% if (lcase(strTimeAdjust)="-11") then Response.Write(" selected") %>>-11</option>
      <option Value="-10"<% if (lcase(strTimeAdjust)="-10") then Response.Write(" selected") %>>-10</option>
      <option Value="-9"<% if (lcase(strTimeAdjust)="-9") then Response.Write(" selected") %>>-9</option>
      <option Value="-8"<% if (lcase(strTimeAdjust)="-8") then Response.Write(" selected") %>>-8</option>
      <option Value="-7"<% if (lcase(strTimeAdjust)="-7") then Response.Write(" selected") %>>-7</option>
      <option Value="-6"<% if (lcase(strTimeAdjust)="-6") then Response.Write(" selected") %>>-6</option>
      <option Value="-5"<% if (lcase(strTimeAdjust)="-5") then Response.Write(" selected") %>>-5</option>
      <option Value="-4"<% if (lcase(strTimeAdjust)="-4") then Response.Write(" selected") %>>-4</option>
      <option Value="-3"<% if (lcase(strTimeAdjust)="-3") then Response.Write(" selected") %>>-3</option>
      <option Value="-2"<% if (lcase(strTimeAdjust)="-2") then Response.Write(" selected") %>>-2</option>
      <option Value="-1"<% if (lcase(strTimeAdjust)="-1") then Response.Write(" selected") %>>-1</option>
      <option Value="0"<% if (lcase(strTimeAdjust)="0") then Response.Write(" selected") %>>0</option>
      <option Value="1"<% if (lcase(strTimeAdjust)="1") then Response.Write(" selected") %>>1</option>
      <option Value="2"<% if (lcase(strTimeAdjust)="2") then Response.Write(" selected") %>>2</option>
      <option Value="3"<% if (lcase(strTimeAdjust)="3") then Response.Write(" selected") %>>3</option>
      <option Value="4"<% if (lcase(strTimeAdjust)="4") then Response.Write(" selected") %>>4</option>
      <option Value="5"<% if (lcase(strTimeAdjust)="5") then Response.Write(" selected") %>>5</option>
      <option Value="6"<% if (lcase(strTimeAdjust)="6") then Response.Write(" selected") %>>6</option>
      <option Value="7"<% if (lcase(strTimeAdjust)="7") then Response.Write(" selected") %>>7</option>
      <option Value="8"<% if (lcase(strTimeAdjust)="8") then Response.Write(" selected") %>>8</option>
      <option Value="9"<% if (lcase(strTimeAdjust)="9") then Response.Write(" selected") %>>9</option>
      <option Value="10"<% if (lcase(strTimeAdjust)="10") then Response.Write(" selected") %>>10</option>
      <option Value="11"<% if (lcase(strTimeAdjust)="11") then Response.Write(" selected") %>>11</option>
      <option Value="12"<% if (lcase(strTimeAdjust)="12") then Response.Write(" selected") %>>12</option>
      <option Value="13"<% if (lcase(strTimeAdjust)="13") then Response.Write(" selected") %>>13</option>
      <option Value="14"<% if (lcase(strTimeAdjust)="14") then Response.Write(" selected") %>>14</option>
      <option Value="15"<% if (lcase(strTimeAdjust)="15") then Response.Write(" selected") %>>15</option>
      <option Value="16"<% if (lcase(strTimeAdjust)="16") then Response.Write(" selected") %>>16</option>
      <option Value="17"<% if (lcase(strTimeAdjust)="17") then Response.Write(" selected") %>>17</option>
      <option Value="18"<% if (lcase(strTimeAdjust)="18") then Response.Write(" selected") %>>18</option>
      <option Value="19"<% if (lcase(strTimeAdjust)="19") then Response.Write(" selected") %>>19</option>
      <option Value="20"<% if (lcase(strTimeAdjust)="20") then Response.Write(" selected") %>>20</option>
      <option Value="21"<% if (lcase(strTimeAdjust)="21") then Response.Write(" selected") %>>21</option>
      <option Value="22"<% if (lcase(strTimeAdjust)="22") then Response.Write(" selected") %>>22</option>
      <option Value="23"<% if (lcase(strTimeAdjust)="23") then Response.Write(" selected") %>>23</option>
      <option Value="24"<% if (lcase(strTimeAdjust)="24") then Response.Write(" selected") %>>24</option>
    </select>
    <a href="JavaScript:openWindow3('pop_config_help.asp#TimeAdjust')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </font></td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">日期格式：</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    <select name="strDateType">
      <option Value="mdy"<% if (lcase(strDateType)="mdy") then Response.Write(" selected") %>>12/31/2000 (US short)</option>
      <option Value="dmy"<% if (lcase(strDateType)="dmy") then Response.Write(" selected") %>>31/12/2000 (UK short)</option>
      <option Value="ymd"<% if (lcase(strDateType)="ymd") then Response.Write(" selected") %>>2000/12/31 (Other short)</option>
      <option Value="ydm"<% if (lcase(strDateType)="ydm") then Response.Write(" selected") %>>2000/31/12 (Other short)</option>
      <option Value="mmdy"<% if (lcase(strDateType)="mmdy") then Response.Write(" selected") %>>Dec 31 2000 (US med)</option>
      <option Value="dmmy"<% if (lcase(strDateType)="dmmy") then Response.Write(" selected") %>>31 Dec 2000 (UK med)</option>
      <option Value="ymmd"<% if (lcase(strDateType)="ymmd") then Response.Write(" selected") %>>2000 Dec 31 (Other med)</option>
      <option Value="ydmm"<% if (lcase(strDateType)="ydmm") then Response.Write(" selected") %>>2000 31 Dec (Other med)</option>
      <option Value="mmmdy"<% if (lcase(strDateType)="mmmdy") then Response.Write(" selected") %>>December 31 2000 (US long)</option>
      <option Value="dmmmy"<% if (lcase(strDateType)="dmmmy") then Response.Write(" selected") %>>31 December 2000 (UK long)</option>
      <option Value="ymmmd"<% if (lcase(strDateType)="ymmmd") then Response.Write(" selected") %>>2000 December 31 (Other long)</option>
      <option Value="ydmmm"<% if (lcase(strDateType)="ydmmm") then Response.Write(" selected") %>>2000 31 December (Other long)</option>
    </select>
    <a href="JavaScript:openWindow3('pop_config_help.asp#datetype')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
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
