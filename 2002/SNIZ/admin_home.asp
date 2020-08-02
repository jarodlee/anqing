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

<table border="0" width="100%">
  <tr>
    <td width="33%" align="left" nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    <img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="default.asp">返回论坛首页</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" height=15 width=15 border="0">&nbsp;论坛管理中心<br>
    </font></td>
  </tr>
</table>

<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td bgcolor="<% =strTableBorderColor %>">
    <table border="0" width="100%" cellspacing="1" cellpadding="3">
      <tr>
        <td bgcolor="<% =strCategoryCellColor %>" colspan=2><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" size"<% =strDefaultFontSize %>"><b>管理设置</b></font></td>
      </tr>
      <tr>
        <td bgcolor="<% =strForumCellColor %>" valign=top><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
        <p><b>论坛功能设定：</b>
        <UL>
        <LI><a href="admin_config_system.asp">基本设置</a></LI>
        <LI><a href="admin_config_features.asp">论坛设置</a></LI>
<%	If strAuthType = "nt" then %>
        <LI><a href="admin_config_NT_features.asp">NT功能设定</a></LI>
<%	End if %>
        <LI><a href="admin_config_members.asp">会员注册选项设置</a></LI>
        <LI><a href="admin_config_ranks.asp">等级设定</a></LI>
        <LI><a href="admin_config_datetime.asp">服务器日期/时间设定</a></LI>
        <LI><a href="admin_config_email.asp">邮件发送设置</a></LI>
        <LI><a href="admin_config_colors.asp">风格设置</a></LI>
        <LI><a href="admin_config_badwords.asp">不良词语过滤</a></LI>
        </UL></p>
        </font></td>
        <td bgcolor="<% =strForumCellColor %>" valign=top><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
        <p><b>其他设定选项与功能：</b>
        <UL>
        <LI><a href="admin_delete_topics.asp">会员贴子管理中心</a></LI>
        <LI><a href="admin_smiles.asp">表情设置</a></LI>
        <LI><a href="admin_moderators.asp">斑竹设置</a></LI>
        <LI><a href="admin_faq.asp">常见问题管理</a></LI>
        <LI><a href="admin_emaillist.asp">邮件列表设置</a></LI>
        <LI><a href="admin_info.asp">服务器环境资料</a></LI>
        <LI><a href="admin_variable_info.asp">论坛系统资讯</a></LI>
        <LI><a href="admin_count.asp">更新论坛状态统计</a></LI>
        <LI><a href="setup.asp">检查安装</a><font size="-2"><b>（每次更新完后请执行本功能！）</b></font></LI>
        </UL></p>
        </font></td>
      </tr>
    </table>
    </td>
  </tr>
</table>
<!--#INCLUDE FILE="mods/modCMD.asp" -->
<!--#INCLUDE file="inc_footer.asp" -->
<% Else %>
<% Response.Redirect "admin_login.asp" %>
<% End IF %>
