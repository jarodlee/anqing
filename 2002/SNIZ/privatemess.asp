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
<%      ' Get Private Message count for display on Default.asp
	if strDBType = "access" then
		strSqL = "SELECT count(M_TO) as [pmcount] " 
	else
        	strSqL = "SELECT count(M_TO) AS pmcount " 
    end if
		strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS , " & strTablePrefix & "PM "
		strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & Request.Cookies(strUniqueID & "User")("Name") & "'"
		strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "PM.M_TO "
		strSql = strSql & " AND " & strTablePrefix & "PM.M_READ = 0 " 
	
		Set rsPM = my_Conn.Execute(strSql)
		pmcount = rsPM("pmcount")
%>
<table border=0 width="80%" cellspacing=1 cellpadding=2 bgcolor="<% =strTableBorderColor %>" >
        <tr>
          <td bgcolor="<% =strCategoryCellColor %>" colspan="2">
            <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" size="+1"><b>���Ļ�ѶϢ</b></font></td>
        </tr>
        <TR>
<td align="center" bgcolor="<% =strForumCellColor %>" valign="middle">
<% 	if Request.Cookies(strUniqueID & "User")("Name") = "" Then %>
		<img src="<%=strImageURL %>icon_pmdead.gif">
<%	else  
 		if pmcount = 0 then %>
			<img src="<%=strImageURL %>icon_pm.gif">
<%		end if 
 		if pmcount >= 1 then %>
			<img src="<%=strImageURL %>icon_pm_new.gif">
<%		end if %>
<%  end if %>
</td>
<TD valign=top bgcolor="<% =strForumCellColor %>" ><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>"><A HREF="pm_view.asp">�ռ���</A></FONT><BR>
<% if Request.Cookies(strUniqueID & "User")("Name") = "" Then %>
<font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>" color="<% =strForumFontColor %>">
    - ��½��̳���ܲ鿴���Ļ���Ϣ
    </font>

<%	else %>
<font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>" color="<% =strForumFontColor %>">
    <b><% if strAuthType="nt" then %>
            <% =session("username")%>&nbsp;(<% =session("userid") %>)</b> 
            <%	else %>	
<%		if strAuthType = "db" then %>
            <% =Request.Cookies(strUniqueID & "User")("Name") %></b><% end if %><% end if %> - ���� <% =pmcount %> ���µ����Ļ�ѶϢ��
    </font>
<%	end if %>
</TD>
</TR>
</table>
