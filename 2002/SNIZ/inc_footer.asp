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
    </div> 
    <td>
    <table width=100% border=0  cellpadding="0" cellspacing = "4"> 
      <tr >
        <td bgcolor="">
        <table border=0 width="100%" align="center" cellpadding="1" cellspacing="0">
          <tr>
            <td bgcolor="<% =strForumCellColor %>" align=center valign=top nowrap><% =strCopyright %></td>
			<td bgcolor="<% =strForumCellColor %>" width=10 nowrap><a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" height=15 width=15 border="0" align="right" alt="顶部"></a></font></td>    
          </tr>
        </table>
        </td>
      </tr>
    </table>
    </td>
  </tr>
  <tr>
	<td>
	  <table border=0 width="100%" align="center" cellpadding="0" cellspacing="0">
	  <tr>
        <td align="left">
        <acronym title="<% =strVersion %> 汉化修改:资源搜罗站"><% if strShowImagePoweredBy = "1" then %><% else %><a href="http://jarod.y365.com" target=_blank>极速伪装空间修改支持</a><% end if %></acronym>
        <td align="right">
        <acronym title="<% =strVersion %>简体中文版"><% if strShowImagePoweredBy = "1" then %><a href="http://jarod.y365.com" target=_blank>极速伪装空间修改支持</a><% else %><Script Src="http://down.my0754.com/logo.js"></Script><% end if %></acronym>
        </td>
      </tr>
	</table>
	</td>
  </tr>
</table>



</font>

</BODY>
</HTML>
<% my_Conn.Close %>
<% set my_Conn = nothing %>
