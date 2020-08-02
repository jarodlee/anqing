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
<%
    HidePM = "HidePM"
       	If Request.QueryString(HidePM) = "Y" then
       		Response.Cookies(HidePM) = "Y"
       		Response.Cookies(HidePM).Expires = dateAdd("d", 30, strForumTimeAdjust)
       	Else
       		If Request.QueryString(HidePM) = "N" then
       			Response.Cookies(HidePM) = "N"
       			Response.Cookies(HidePM).Expires = dateAdd("d", 30, strForumTimeAdjust)
      		End If
        End If
        
        ' Get Private Message count for display on Default.asp
	if strDBType = "access" then
		strSqL = "SELECT count(M_TO) as [pmcount] " 
	else
        	strSQL = "SELECT count(FORUM_ID) forums " 
    	end if
		strSql = strSql & " FROM " & strTablePrefix & "MEMBERS , " & strTablePrefix & "PM "
		strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & Request.Cookies(strUniqueID & "User")("Name") & "'"
		strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "PM.M_TO "
		strSql = strSql & " AND " & strTablePrefix & "PM.M_READ = 0 " 
	
		Set rsPM = my_Conn.Execute(strSql)
		pmcount = rsPM("pmcount")
%>

        <tr>
          <td bgcolor="<% =strHeadCellColor %>" colspan="<% if (strShowModerators = "1") or (mlev = 4 or mlev = 3) then Response.Write("7") else Response.Write("5")%>">
<%		' This code will specify whether or not to show the forums under a category
		HidePM = "HidePM"
 		If Request.Cookies(HidePM) = "Y" then %>
        		<a href="<% =ScriptName & "?" %><% =HidePM & "=N" %>"><img src="<%=strImageURL %>plus.gif" width="10" height="10" border="0" alt="Expand Private Messages"></a>
<% 		Else %>
       			<a href="<% =ScriptName & "?" %><% =HidePM & "=Y" %>"><img src="<%=strImageURL %>minus.gif" width="10" height="10" border="0" alt="Collapse Private Messages"></a>
<% 		End If  %>
            <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" size="+1"><b>悄悄话讯息</b></font></td>
        </tr>
<%     if Request.Cookies(HidePM) <> "Y" then %>
        <TR>
<td align="center" bgcolor="<% =strForumCellColor %>" valign="middle">
<% if Request.Cookies(strUniqueID & "User")("Name") = "" Then %>
<img src="<%=strImageURL %>pmdead.gif">
<%	else  

 if pmcount = 0 then %>
<img src="<%=strImageURL %>icon_folder.gif">
<%	end if 
 if pmcount >= 1 then %>
<img src="<%=strImageURL %>icon_folder_new.gif">
<%	end if %>
<%  end if %>
</td>
<TD valign=top bgcolor="<% =strForumCellColor %>" colspan="<% if (strShowModerators = "1") or (mlev = 4 or mlev = 3) then Response.Write("7") else Response.Write("5")%>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" size="+1"><A HREF="pm_view.asp">悄悄话收件箱</A></FONT><BR>
<% if Request.Cookies(strUniqueID & "User")("Name") = "" Then %>
<font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>" color="<% =strActiveLinkColor %>">
    - 登陆论坛才能查看悄悄话信息
    </font>

<%	else %>
<font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>" color="<% =strActiveLinkColor %>">
    <b><% if strAuthType="nt" then %>
            <% =session("username")%>&nbsp;(<% =session("userid") %>)</b> 
            <%	else %>	
<%		if strAuthType = "db" then %>
            <% =Request.Cookies(strUniqueID & "User")("Name") %></b><% end if %><% end if %> - 你有 <% =pmcount %> 条新的悄悄话讯息。
    </font>
    
</TD>
</TR>
<%	end if 
end if %>