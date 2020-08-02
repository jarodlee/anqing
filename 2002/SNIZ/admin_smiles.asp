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
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<br>
<%
	intpageSize = 10
	
	set drs = server.CreateObject("ADODB.RecordSet")
	
	drs.CacheSize= intpageSize
	strsql = "select * from " & strTablePrefix & "smiles"
    drs.Open strsql, my_Conn,3
	iPage = CInt(request("iPage"))
	drs.PageSize = intpageSize
	
	iPageCount = cInt(drs.PageCount)

	if iPage < 1 then
		iPage = 1
	end if
	drs.AbsolutePage = iPage
    rec = 0
%>

<table border=0 align="center" cellspacing="1" cellpadding="4">
<tr><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
<p align="center">插入新表情图标之前请确定图标图片已放在表情图标存放的文件夹里面</p>
<div align="center"><a href="admin_home.asp">返回论坛管理中心</a></div>
</font>
<br>
</tr>
<tr><th bgcolor="<% =strHeadCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<%=strHeadFontColor%>">表情设置</font></th></tr>
<tr><td valign="top">
<table border=0 align="center" bgcolor="<% =strTableBorderColor %>" cellspacing="1" cellpadding="4">
<tr bgcolor="<% =strForumCellColor %>" ><td colspan="7">
<%
							if iPage > 1 then
								response.write "&nbsp;&nbsp;&nbsp;<a href=""admin_smiles.asp?iPage=1"" alt=""第一页"">[第一页]&nbsp;&nbsp;&nbsp;</a>"
								response.write "<a href=""admin_smiles.asp?iPage=" & iPage - 1 & """ alt=""上一页"">"
								response.write "上一页</a>&nbsp;|&nbsp;"
							else
								response.write "&nbsp;&nbsp;&nbsp;[第一页]&nbsp;&nbsp;&nbsp;上一页&nbsp;|&nbsp;"
							end if
							if iPage < iPageCount then 
								response.write "<a href=""admin_smiles.asp?iPage=" & iPage + 1 & """ alt=""下一页"">"
								response.write "下一页</a>"
								response.write "&nbsp;&nbsp;&nbsp;<a href=""admin_smiles.asp?iPage=" & iPageCount & """ alt=""最后一页"">[最后一页]</a>"
							else
								response.write "下一页&nbsp;&nbsp;&nbsp;[最后一页]"
							end if
							response.write "&nbsp;&nbsp;&nbsp;[页数：" & iPage & "/" & iPageCount & "]"
%>
</td></tr>
<tr bgcolor="<% =strForumCellColor %>"><th><font face="<%=strDefaultFontFace%>" size=<%=strDefaultFontSize%>>分类</font></th><th><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">ID</font></th><th><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">图示</font></th><th><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">代码</font></th><th><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">表情图标地址</font></th><th><font face="<%=strDefaultFontFace%>" size="<%=strDefaultFontSize%>">名称</font></th><th><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">动作</font></th></tr>
<%	Do until drs.eof or rec = intPageSize
    	smile_id = drs("id")
        smile_code = drs("smile_code")
        smile_url = drs("smile_url")
        smile_name = drs("smile_name")
        smile_cat = drs("cat_id")
                %>
        <form method="post" action="edit_smile.asp">
        <input type="hidden" name="smile_id" value="<%=smile_id%>">
        <tr bgcolor="<% =strForumCellColor %>">
        <td>
        <font face="<%=strDefaultFontFace %>" size="<%=strDefaultFontSize+1%>">
        <select name="cat_id" size=1>
	    <%
        strsql = "select cat_id, cat_name from "&strTablePrefix&"smile_cat"
        set catRs = my_conn.execute(strsql)
        	Do until catRS.eof
            %>
        <option value="<%=catRs("cat_id")%>"<%if catRs("cat_id") = smile_cat then%> SELECTED<%end if%>><%=catRS("cat_name")%></option>
        <%
        catRs.movenext
        loop
        set catRs = nothing
        %>
        </select
        </font>
        </td>
        
        <td><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize+1 %>"><%=smile_id%></font></td><td bgcolor="<% =strForumCellColor %>" align="center"><img src="<%=strImageURL %><%=smile_url%>" alt="<%=smile_name%>"></td><td>[<input type="text" name="smile_code" maxlength="10" size="5" value="<%=smile_code%>">]</td><td><input type="text" name="smile_url" maxlength="50" value="<%=smile_url%>"></td>
<td><input type="text" name="smile_name" value="<%=smile_name%>" size="15" maxlength="40"></td>
<td><input type="submit" value="更新"> <a href="delete_smile.asp?smile_id=<%=smile_id%>" onclick="return confirm('确定删除此表情符号？')"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">删除</font></a></td></tr></form>
        <%
	rec = rec + 1
    drs.movenext
    loop
    set drs=nothing
%>
</table>
</td>
</tr>
<tr><th bgcolor="<% =strHeadCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<%=strHeadFontColor%>">新增表情符号</font></th></tr>
<tr><td>

<table border=0 cellspacing="1" bgcolor="<% =strTableBorderColor %>" cellpadding="3" width="100%">
<tr bgcolor="<% =strForumCellColor %>">
<th><font face="<% = strDefaultFontFace %>" size="<% =strDefaultFontSize %>">分类</font></th>
<th><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">代码</font></th><th><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">表情图标地址</font></th><th><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">名称</font></th><th>&nbsp;</th></tr>
<form action="add_smile.asp" method="post">
<tr bgcolor="<% =strForumCellColor %>">
	<td><select name="cat_id" size="1">
    <%
        strsql = "select cat_id, cat_name from "&strTablePrefix&"smile_cat"
        set catRs = my_conn.execute(strsql)
        	Do until catRS.eof
            %>
        <option value="<%=catRs("cat_id")%>"><%=catRs("cat_name")%></option>
        <%
        catRs.movenext
        loop
        set catRs = nothing
        %>
        </select></td>
        <td>[<input type="text" name="smile_code" size="5" maxlength="10">]</td>
	<td><input type="text" name="smile_url" size="20" maxlength="25"></td>
    <td><input type="text" name="smile_name" size="15" maxlength="20"></td>
    <td><input type="submit" Value="新增"></td>
</tr>
</table>
</form>
<form method="post" action="add_smile.asp?type=cat">
<table border=0 cellspacing="1" bgcolor="<% =strTableBorderColor %>" cellpadding="3" width="100%">
<tr bgcolor="<% =strForumCellColor %>">
	<th bgcolor="<%=strHeadCellColor%>" ><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<%=strHeadFontColor%>">新增分类</font></th>
</tr>
<tr BGCOLOR="<%=strForumCellColor%>">
	<td><input type="text" maxlength="30" name="cat_name">&nbsp;<input type="submit" value="新增分类"></td>
</tr>
</table>
</form>
<form method="post" action="delete_smile.asp?type=cat">
<table border=0 cellspacing=1 bgcolor="<%=strTableBorderColor%>" cellpadding=1 width=100%
	<tr bgcolor="<%=strForumCellColor%>">
    <th bgcolor="<%=strHeadCellColor%>"><font color="<%=strHeadFontColor%>" size=<%=strDefaultFontSize%>>删除分类</font></th>
    </tr>
    <tr bgcolor="<%=strForumCellColor%>">
    <td>
<select name="cat_id" size="1">
    <%
        strsql = "select cat_id, cat_name from "&strTablePrefix&"smile_cat"
        set catRs = my_conn.execute(strsql)
        	Do until catRS.eof
            %>
        <option value="<%=catRs("cat_id")%>"><%=catRs("cat_name")%></option>
        <%
        catRs.movenext
        loop
        set catRs = nothing
        %>
        </select>
        <input type="submit" value="删除" onclick="return confirm('删除此分类及所包含的表情符号？')">
        </td>
      </tr>
      </table>
      </form>
</td></tr></table>
<!--#INCLUDE file="inc_footer.asp" -->
<%else%>
你不是管理者！
<!--#include file="inc_footer.asp" -->
<%end if%>
