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
<p align="center">�����±���ͼ��֮ǰ��ȷ��ͼ��ͼƬ�ѷ��ڱ���ͼ���ŵ��ļ�������</p>
<div align="center"><a href="admin_home.asp">������̳��������</a></div>
</font>
<br>
</tr>
<tr><th bgcolor="<% =strHeadCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<%=strHeadFontColor%>">��������</font></th></tr>
<tr><td valign="top">
<table border=0 align="center" bgcolor="<% =strTableBorderColor %>" cellspacing="1" cellpadding="4">
<tr bgcolor="<% =strForumCellColor %>" ><td colspan="7">
<%
							if iPage > 1 then
								response.write "&nbsp;&nbsp;&nbsp;<a href=""admin_smiles.asp?iPage=1"" alt=""��һҳ"">[��һҳ]&nbsp;&nbsp;&nbsp;</a>"
								response.write "<a href=""admin_smiles.asp?iPage=" & iPage - 1 & """ alt=""��һҳ"">"
								response.write "��һҳ</a>&nbsp;|&nbsp;"
							else
								response.write "&nbsp;&nbsp;&nbsp;[��һҳ]&nbsp;&nbsp;&nbsp;��һҳ&nbsp;|&nbsp;"
							end if
							if iPage < iPageCount then 
								response.write "<a href=""admin_smiles.asp?iPage=" & iPage + 1 & """ alt=""��һҳ"">"
								response.write "��һҳ</a>"
								response.write "&nbsp;&nbsp;&nbsp;<a href=""admin_smiles.asp?iPage=" & iPageCount & """ alt=""���һҳ"">[���һҳ]</a>"
							else
								response.write "��һҳ&nbsp;&nbsp;&nbsp;[���һҳ]"
							end if
							response.write "&nbsp;&nbsp;&nbsp;[ҳ����" & iPage & "/" & iPageCount & "]"
%>
</td></tr>
<tr bgcolor="<% =strForumCellColor %>"><th><font face="<%=strDefaultFontFace%>" size=<%=strDefaultFontSize%>>����</font></th><th><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">ID</font></th><th><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">ͼʾ</font></th><th><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">����</font></th><th><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">����ͼ���ַ</font></th><th><font face="<%=strDefaultFontFace%>" size="<%=strDefaultFontSize%>">����</font></th><th><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">����</font></th></tr>
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
<td><input type="submit" value="����"> <a href="delete_smile.asp?smile_id=<%=smile_id%>" onclick="return confirm('ȷ��ɾ���˱�����ţ�')"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">ɾ��</font></a></td></tr></form>
        <%
	rec = rec + 1
    drs.movenext
    loop
    set drs=nothing
%>
</table>
</td>
</tr>
<tr><th bgcolor="<% =strHeadCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<%=strHeadFontColor%>">�����������</font></th></tr>
<tr><td>

<table border=0 cellspacing="1" bgcolor="<% =strTableBorderColor %>" cellpadding="3" width="100%">
<tr bgcolor="<% =strForumCellColor %>">
<th><font face="<% = strDefaultFontFace %>" size="<% =strDefaultFontSize %>">����</font></th>
<th><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">����</font></th><th><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">����ͼ���ַ</font></th><th><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">����</font></th><th>&nbsp;</th></tr>
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
    <td><input type="submit" Value="����"></td>
</tr>
</table>
</form>
<form method="post" action="add_smile.asp?type=cat">
<table border=0 cellspacing="1" bgcolor="<% =strTableBorderColor %>" cellpadding="3" width="100%">
<tr bgcolor="<% =strForumCellColor %>">
	<th bgcolor="<%=strHeadCellColor%>" ><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<%=strHeadFontColor%>">��������</font></th>
</tr>
<tr BGCOLOR="<%=strForumCellColor%>">
	<td><input type="text" maxlength="30" name="cat_name">&nbsp;<input type="submit" value="��������"></td>
</tr>
</table>
</form>
<form method="post" action="delete_smile.asp?type=cat">
<table border=0 cellspacing=1 bgcolor="<%=strTableBorderColor%>" cellpadding=1 width=100%
	<tr bgcolor="<%=strForumCellColor%>">
    <th bgcolor="<%=strHeadCellColor%>"><font color="<%=strHeadFontColor%>" size=<%=strDefaultFontSize%>>ɾ������</font></th>
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
        <input type="submit" value="ɾ��" onclick="return confirm('ɾ���˷��༰�������ı�����ţ�')">
        </td>
      </tr>
      </table>
      </form>
</td></tr></table>
<!--#INCLUDE file="inc_footer.asp" -->
<%else%>
�㲻�ǹ����ߣ�
<!--#include file="inc_footer.asp" -->
<%end if%>
