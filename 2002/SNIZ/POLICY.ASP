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
<!--#INCLUDE FILE="inc_top.asp" -->

<table width="100%" border=0>
  <tr>
    <td>
    <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><img src="<%=strImageURL %>icon_folder_open.gif" border="0">&nbsp;<a href="default.asp">返回论坛首页</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" border="0">&nbsp;服务条款和声明</font>
    </td>
  </tr>
</table>
<table border="0" width="95%" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td bgcolor="<% =strTableBorderColor %>">
<table border="0" width="100%" cellspacing="1" cellpadding="4">
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b><% =strForumTitle %>服务条款和声明</b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
<p>如果你同意以下所列的条款，请按本页最下方的<b>【我同意】</b>钮，如果不同意，请按<b>【我不同意】</b>钮。

</p>

<p>您在本站注册所填写的个人资料，均受「电脑处理个人资料保密法」所保　，本站依法不会将会员资料提供其他公司或第三者使用。惟基於管理　要，本站於每个讨论区设有板主一职，板主在职权范围内可视　求检阅用户所填写的资料。</p>

<p>用户在本站的留言其著作权均归其个人所有，但每则留均表示授权本站可在网站上自由使用之。由於我们对这个站台具有管理之责，因此你也同意本站得视　要对使用者的留言进行修改、删除、搬移等之必要维护。本站也保留取消不受欢迎的用户的会员资格权利。此外，本站适用国内之一切法律，请勿在本站刊载违法之图文，或交流色情、盗版等资讯，如果有违规之情事，本站将予以删除。</p>

<p>本站各版之细部管理规则由各板之板主自行订定之。用各於各板均应遵守各板板主所订定之规则，如果板主未订定规则者，则以本站之一般性通用规则为准。除此之外，用户留言应注意以下事项：

<ol>
  <li>请确实依各板主题性质留言，避免文不对题。
  <li><font color=red>本站严禁张贴广告，<b>特别是网路赚钱广告</b>。如果有违规立即取消会员资格！</font>
  <li>相同的主题请勿跨板张贴, 以节省网路资源。 
  <li>如有任何问题，欢迎来信询问
  <a href="mailto:<% =strSender %>"><% =strSender %></a>。</li>
</ol>
  
</p>

<p>感谢您的配合，祝您使用愉快！</p>

<hr size=1>

<table align="center" border="0">
  <tbody>
    <tr>
      <td>
        <form action="register.asp?mode=Register" id="FORM1" method="post" name="FORM1">
          <input name="Refer" type="hidden" value="<% =Request.ServerVariables("HTTP_REFERER") %>">
          <input name="submit" type="submit" value="同意">
        </form>
      </td>
      <td>
        <form action="JavaScript:history.go(-1)" id="FORM2" method="post" name="FORM2">
          <input name="submit" type="submit" value="离开">
        </form>
      </td>
    </tr>
  </tbody>
</table>
      </font></td>
  </tr>
</table>
    </td>
  </tr>
</table>
<!--#INCLUDE FILE="inc_footer.asp" -->
