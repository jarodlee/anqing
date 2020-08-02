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
            <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" size="+1"><b>悄悄话讯息</b></font></td>
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
<TD valign=top bgcolor="<% =strForumCellColor %>" ><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>"><A HREF="pm_view.asp">收件箱</A></FONT><BR>
<% if Request.Cookies(strUniqueID & "User")("Name") = "" Then %>
<font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>" color="<% =strForumFontColor %>">
    - 登陆论坛才能查看悄悄话信息
    </font>

<%	else %>
<font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>" color="<% =strForumFontColor %>">
    <b><% if strAuthType="nt" then %>
            <% =session("username")%>&nbsp;(<% =session("userid") %>)</b> 
            <%	else %>	
<%		if strAuthType = "db" then %>
            <% =Request.Cookies(strUniqueID & "User")("Name") %></b><% end if %><% end if %> - 你有 <% =pmcount %> 条新的悄悄话讯息。
    </font>
<%	end if %>
</TD>
</TR>
</table>
