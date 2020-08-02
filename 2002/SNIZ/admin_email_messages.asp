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
<%
strsql = "Select * from " & strTablePrefix & "email_config where m_type = 'newreply'"

set drs = my_conn.execute(strsql)
	strTopicAuthorSubject = drs("m_subject")
    strTopicAuthorMessage = drs("m_message")


strsql = "Select * from "&strTablePrefix&"email_config where "&strTablePrefix&"email_config.m_type='replynot'"

set drs = my_conn.execute(strsql)
	strReplyNotSubject = drs("m_subject")
    strReplyNotMessage = drs("m_message")

%>
<table border="0" width="100%" >
  <tr>
    <td width="33%" align="left" nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    <img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="default.asp">返回论坛首页</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="admin_home.asp">论坛管理中心</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" height=15 width=15 border="0">&nbsp;邮件信息设置
    </font></td>
  </tr>
</table><br>
<table border=0 cellspacing=1 cellpadding=3 align="center" bgcolor="<%=strTableBorderColor%>" width="100%" >
	<tr><td bgcolor="<% =strHeadCellColor %>" colspan="4"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">邮件信息设置</font>
</td></tr>
	<tr>
    	<td bgcolor="<%=strForumCellColor%>" align="left" colspan="4">
        下面是论坛名称、文章主题等等的快捷代码列表:<br></td></tr>
        <tr><td bgcolor="<%=strForumCellColor%>" width="25%">[forum_title]</td><td bgcolor="<%=strForumCellColor%>" colspan="3">论坛名称.</td></tr>
        <tr><td bgcolor="<%=strForumCellColor%>">[user_name]</td><td bgcolor="<%=strForumCellColor%>" colspan="3">会员名字.</td></tr>
        <tr><td bgcolor="<%=strForumCellColor%>">[topic_title]</td><td bgcolor="<%=strForumCellColor%>" colspan="3">贴子主题.</td></tr>
        <tr><td bgcolor="<%=strForumCellColor%>">[link]</td><td bgcolor="<%=strForumCellColor%>" colspan="3">主题连接地址.</td></tr>
        <tr><td bgcolor="<%=strForumCellColor%>">[posted_by]</td><td bgcolor="<%=strForumCellColor%>" colspan="3">主题作者</td></tr>
		<tr><td colspan=4 bgcolor="<%=strForumCellColor%>"><font color="red"><b>注意:</b> [forum_title] 只能放在信件标题中..</font></td></tr>
    </tr>
</table>
<br>
<table border=0 cellspacing=1 cellpadding=3 align="center" bgcolor="<%=strTableBorderColor%>" width="100%" >
	<tr>
		<td bgcolor="<%=strHeadCellColor%>" width="50%" ><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>"><b>主题作者回复通知</b></font></td>
		<td bgcolor="<%=strHeadCellColor%>" width="50%"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>"><b>Requested Notification</b></font></td>				
    </tr>
    <tr>
    	<td bgcolor="<%=strForumCellColor%>">当有回复时通过email通知作者.</td>
		<td bgcolor="<%=strForumCellColor%>">这个是给那些回复这个主题的人的。并向他要通知</td>
	</tr>
<td bgcolor="<%=strForumCellColor%>">
		<table>
        <form method="post" action="post_email_config.asp">
    	<tr>
      		<th align="left" valign="top"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strDefaultFontColor %>">标题：</font></th>
      		<td><input type="text" name="tsubject" size="40" maxlength="100" value="<%=strTopicAuthorSubject%>"></td>
    	</tr>
    	<tr>
      		<th align="left" valign="top"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strDefaultFontColor %>">内容：</font></th>
      		<td><textarea name="tmessage" cols="45" rows="6" wrap="virtual"><%=strTopicAuthorMessage%></textarea></td>
    	</tr>
    	<tr>
    		<td colspan=2 align="right"><input type="submit" value=" 修 改 "></td>
    	</tr>
		</table>
        </td>
		<td bgcolor="<%=strForumCellColor%>">
			<table>

		    <tr>
		      <th align="left" valign="top">标题：</th>
		      <td><input type="text" name="rsubject" size="40" maxlength="100" value="<%=strReplyNotSubject%>"></td>
		    </tr>
		    <tr>
		      <th align="left" valign="top">内容：</th>
		      <td><textarea name="rmessage" cols="45" rows="6" wrap="virtual"><%=strReplyNotMessage%></textarea></td>
		    </tr>
		    <tr>
		    	<td colspan=2 align="right"><input type="submit" value=" 修 改 "></td>
		    </tr>
			</form>
			</table>
		</td>
    </tr>
</table>




<!--#INCLUDE file="inc_footer.asp" -->
<% Else %>
<% Response.Redirect "admin_login.asp" %>
<% End IF %>
