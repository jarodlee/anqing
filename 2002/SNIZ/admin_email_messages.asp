<%
'########## Snitz Forums 2000 Version 3.1 SR4 ####################
'#                                                               #
'#  �����޸�: ��Դ����վ                                         #
'#  �����ʼ�: cgier@21cn.com                                     #
'#  ��ҳ��ַ: http://www.sdsea.com                               #
'#            http://www.99ss.net                                #
'#            http://www.cdown.net                               #
'#	     http://www.wzdown.com                               #
'#	     http://www.13888.net                                #
'#  ��̳��ַ:http://ubb.yesky.net                                #
'#  ����޸�����: 2001/03/12    ���İ汾��Version 3.1 SR4        #
'#################################################################
'# ԭʼ��Դ                                                      #
'# Snitz Forums 2000 Version 3.1 SR4                             #
'# Copyright 2000 http://forum.snitz.com - All Rights Reserved   #
'#################################################################
'#����Ȩ������                                                   #
'#                                                               #
'# ������Ϊ��������(shareware)�ṩ������վ���ʹ�ã�����Ƿ��޸�,#
'# ת�أ�ɢ��������������ͼ����Ϊ��������ɾ����Ȩ������          #
'# ���������վ��ʽ����������ű�������֪ͨ���ǣ��Ա������ܹ�֪��#
'# ������ܣ�����������վ�������ǵ����ӣ�ϣ���ܸ��������лл��  #
'#################################################################
'# �����������ǵ��Ͷ��Ͱ�Ȩ����Ҫɾ�����ϵİ�Ȩ�������֣�лл����#
'# �����κ������뵽���ǵ���̳��������                            #
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
    <img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="default.asp">������̳��ҳ</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="admin_home.asp">��̳��������</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" height=15 width=15 border="0">&nbsp;�ʼ���Ϣ����
    </font></td>
  </tr>
</table><br>
<table border=0 cellspacing=1 cellpadding=3 align="center" bgcolor="<%=strTableBorderColor%>" width="100%" >
	<tr><td bgcolor="<% =strHeadCellColor %>" colspan="4"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">�ʼ���Ϣ����</font>
</td></tr>
	<tr>
    	<td bgcolor="<%=strForumCellColor%>" align="left" colspan="4">
        ��������̳���ơ���������ȵȵĿ�ݴ����б�:<br></td></tr>
        <tr><td bgcolor="<%=strForumCellColor%>" width="25%">[forum_title]</td><td bgcolor="<%=strForumCellColor%>" colspan="3">��̳����.</td></tr>
        <tr><td bgcolor="<%=strForumCellColor%>">[user_name]</td><td bgcolor="<%=strForumCellColor%>" colspan="3">��Ա����.</td></tr>
        <tr><td bgcolor="<%=strForumCellColor%>">[topic_title]</td><td bgcolor="<%=strForumCellColor%>" colspan="3">��������.</td></tr>
        <tr><td bgcolor="<%=strForumCellColor%>">[link]</td><td bgcolor="<%=strForumCellColor%>" colspan="3">�������ӵ�ַ.</td></tr>
        <tr><td bgcolor="<%=strForumCellColor%>">[posted_by]</td><td bgcolor="<%=strForumCellColor%>" colspan="3">��������</td></tr>
		<tr><td colspan=4 bgcolor="<%=strForumCellColor%>"><font color="red"><b>ע��:</b> [forum_title] ֻ�ܷ����ż�������..</font></td></tr>
    </tr>
</table>
<br>
<table border=0 cellspacing=1 cellpadding=3 align="center" bgcolor="<%=strTableBorderColor%>" width="100%" >
	<tr>
		<td bgcolor="<%=strHeadCellColor%>" width="50%" ><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>"><b>�������߻ظ�֪ͨ</b></font></td>
		<td bgcolor="<%=strHeadCellColor%>" width="50%"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>"><b>Requested Notification</b></font></td>				
    </tr>
    <tr>
    	<td bgcolor="<%=strForumCellColor%>">���лظ�ʱͨ��email֪ͨ����.</td>
		<td bgcolor="<%=strForumCellColor%>">����Ǹ���Щ�ظ����������˵ġ�������Ҫ֪ͨ</td>
	</tr>
<td bgcolor="<%=strForumCellColor%>">
		<table>
        <form method="post" action="post_email_config.asp">
    	<tr>
      		<th align="left" valign="top"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strDefaultFontColor %>">���⣺</font></th>
      		<td><input type="text" name="tsubject" size="40" maxlength="100" value="<%=strTopicAuthorSubject%>"></td>
    	</tr>
    	<tr>
      		<th align="left" valign="top"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strDefaultFontColor %>">���ݣ�</font></th>
      		<td><textarea name="tmessage" cols="45" rows="6" wrap="virtual"><%=strTopicAuthorMessage%></textarea></td>
    	</tr>
    	<tr>
    		<td colspan=2 align="right"><input type="submit" value=" �� �� "></td>
    	</tr>
		</table>
        </td>
		<td bgcolor="<%=strForumCellColor%>">
			<table>

		    <tr>
		      <th align="left" valign="top">���⣺</th>
		      <td><input type="text" name="rsubject" size="40" maxlength="100" value="<%=strReplyNotSubject%>"></td>
		    </tr>
		    <tr>
		      <th align="left" valign="top">���ݣ�</th>
		      <td><textarea name="rmessage" cols="45" rows="6" wrap="virtual"><%=strReplyNotMessage%></textarea></td>
		    </tr>
		    <tr>
		    	<td colspan=2 align="right"><input type="submit" value=" �� �� "></td>
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
