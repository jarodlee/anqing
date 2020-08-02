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
<!--#INCLUDE FILE="inc_top_short.asp" -->
<% select case Request.QueryString("mode") %>
<%	case "ICQ" %>
<p><font face="<% =strDefaultFontFace %>"size="<% =strHeaderFontSize %>">发送ICQ讯息</font></p>
<form action="http://wwp.mirabilis.com/scripts/WWPMsg.dll" method="post">
<input type="hidden" name="subject" value="From Your Web Page">
<input type="hidden" name="to" value="<% =Request.QueryString("ICQ") %>">

<table border="0" width="75%" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="<% =strPopUpBorderColor %>">
    <table border="0" width="100%" cellspacing="1" cellpadding="1">
      <TR>
        <TD bgColor=<% =strPopUpTableColor %> align=right nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">收件人：</font></td>
        <TD bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% =Request.QueryString("M_NAME")%></font></td>
      </tr>
      <TR>
        <TD bgColor=<% =strPopUpTableColor %> align=right nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">收件人 ICQ 号码：</font></td>
        <TD bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><img src="http://online.mirabilis.com/scripts/online.dll?icq=<% =Request.QueryString("ICQ") %>&img=5" border="0" align="absmiddle"><% =Request.QueryString("ICQ") %></font></td>
      </tr>
      <TR>
        <TD bgColor=<% =strPopUpTableColor %> align=right nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">发件人姓名：</font></td>
        <TD bgColor=<% =strPopUpTableColor %>><input type="text" name="from" value size="20" maxlength="40" onfocus="this.select()"></td>
      <TR>
        <TD bgColor=<% =strPopUpTableColor %> align=right nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">发件人电子邮件：</FONT></td>
        <TD bgColor=<% =strPopUpTableColor %>><input type="text" name="fromemail" value size="20" maxlength="40" onfocus="this.select()"></td>
      <TR>
        <TD bgColor=<% =strPopUpTableColor %> Colspan=2 align=center nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">讯息</font></td>
      </TR>
      <TR>
        <TD bgColor=<% =strPopUpTableColor %> Colspan=2 align=center><textarea name="body" rows="3" cols="40" wrap="Virtual"></textarea></td>
      </TR>
      <TR>
        <TD bgColor=<% =strPopUpTableColor %> Colspan=2 align=center><input type=submit value="发送"></td>
      </tr>
    </TABLE>
    </td>
  </tr>
</table>
</form>
<%	case "AIM" %>
<p><font face="<% =strDefaultFontFace %>"size="<% =strHeaderFontSize %>"><% =Request.QueryString("M_NAME") %>的AIM设置</font></p>

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>注意：</b> 你必须安装了 AOL Instant Messenger 才能执行此功能。</font></p>
<table border="0" width="75%" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="<% =strPopUpBorderColor %>">
    <table border="0" width="100%" cellspacing="1" cellpadding="1">
      <TR>
        <TD bgColor=<% =strPopUpTableColor %> align=center nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="aim:goIM?screenname=<% =Request.QueryString("AIM") %>" alt="传送讯息">传送讯息</a></font></td>
      </tr>
      <TR>
        <TD bgColor=<% =strPopUpTableColor %> align=center nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="aim:goChat?ROOMname=<% =Request.QueryString("AIM") %>">开启聊天室</a></font></td>
      </tr>
      <TR>
        <TD bgColor=<% =strPopUpTableColor %> align=center nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="aim:addBuddy?screenname=<% =Request.QueryString("AIM") %>">加入好友名单</a></font></td>
      </tr>
    </TABLE>
    </td>
  </tr>
</table>

<% end select %>
<% If InStr(Request.ServerVariables("HTTP_REFERER"), "pop_profile.asp") <> 0 then %>
<p><font size="<% =strDefaultFontSize %>"><a href="JavaScript:history.go(-1)">回到 <% =Request.QueryString("M_NAME")%> 的会员资料</a></font></p>
<% end if %>

<!--#INCLUDE FILE="inc_footer_short.asp" -->
