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

<table border="0" width="95%" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td bgcolor="<% =strTableBorderColor %>">
<table border="0" width="100%" cellspacing="1" cellpadding="4">
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="hyperlink"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>在文章中使用超级连接</b></font></td>
  </tr>
  <tr>
      <td bgcolor="<% =strForumCellColor %>"><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
      <p>你可以在文章中轻松的加入超级连接。</p>
      
      <p>你可以直接输入完整的网址（<% =strForumURL %>），系统会自行转成连接网址（<a href="<% =strForumURL %>" target="_blank"><% =strForumURL %></a>）！</p>
      <p>注意！一定要在网址前加上 <b>http://</b>, <b>https://</b> 或 <b>file://</b></p>

	  <p>另一种使用超级连接的方式是使用下面的 <b>[url]</b>网址<b>[/url]</b> 标签</p>
<!--	  <blockquote>
              <i>范例：</i><br>
              <b>[url]</b><% =strForumURL %><b>[/url]</b> 回到论坛首页！<br>
              <i>会变成：</i><br>
              <a href="<% =strForumURL %>"><% =strForumURL %></a> 回到论坛首页！
      </blockquote></p>
	  <p> -->
      <p>如果你使用以下格式：<b>[url=&quot;</b>连接<b>&quot;]</b>描述<b>[/url]</b> 你可以加入一个连接的名称或描述。</p>
<!--      <blockquote>
              <i>示例：</i><br>
              连接到 <b>[url=&quot;<% =strForumURL %>&quot;]</b><% =strForumTitle %><b>[/url]</b><br>
              <i>会变成：</i><br>
              连接到 <a href="<% =strForumURL %>"><% =strForumTitle %></a>
      </blockquote></p> -->
      </font></td>
  </tr>
</table>
    </td>
  </tr>
</table>

<!--#INCLUDE FILE="inc_footer_short.asp" -->
