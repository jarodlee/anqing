<!--#include file=conn.asp-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta NAME=GENERATOR Content="Microsoft FrontPage 3.0" CHARSET=GB2312>
<BODY leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<%
dim admin_flag
admin_flag="17"
if not master or instr(session("flag"),admin_flag)=0 then
		Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��½</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		call dvbbs_error()
	else
		call main()
		conn.close
		set conn=nothing
	end if

	sub main()
if request("action")="save" then
call savegroup()
elseif request("action")="savedit" then
call savedit()
elseif request("action")="del" then
call del()
else
call gradeinfo()
end if
end sub

sub gradeinfo()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<th width="100%" class="tableHeaderText" colspan=2>��̳���ɹ���
</th>
</tr>
<tr>
<td height="23" colspan="2" class=forumrow><b>�û����ɹ���</b>������������޸Ļ���ɾ����̳���ɡ�</td>
</tr>
<%if request("action")="edit" then%>
<form method="POST" action=admin_menpai.asp?action=savedit>
<%
	set rs=conn.execute("select * from GroupName where id="&request("id"))
%>
<tr> 
<th width="100%" id=TableTitleLink colspan=2>�޸����� | <a href=admin_menpai.asp><b>�������<b></a></th>
</tr>
<tr> 
<td width="30%" class=forumrow><b>��������</b></td>
<td width="70%" class=forumrow> 
<input type="text" name="Groupname" size="35" value="<%=rs("Groupname")%>">&nbsp;<input type="submit" name="Submit" value="�� ��">
<input type=hidden name=id value="<%=request("id")%>">
</td>
</tr>
<%set rs=nothing%>
<%else%>
<form method="POST" action=admin_menpai.asp?action=save>
<tr> 
<th width="100%" class="tableHeaderText" colspan=2><b>�������</b>
</th>
</tr>
<tr> 
<td width="30%" class=forumrow><b>����</B>����</td>
<td width="70%" class=forumrow>
<input type="text" name="Groupname" size="35">&nbsp;<input type="submit" name="Submit" value="�� ��">
</td>
</tr>
</form>
<%end if%>
<tr> 
<th height="23" colspan="2" ><b>��������</b></th>
</tr>
<%
	set rs=conn.execute("select * from GroupName")
	do while not rs.eof
%>
<tr> 
<td height="23" colspan="2" class=forumRowHighlight>
<a href="admin_menpai.asp?id=<%=rs("id")%>&action=edit">�޸�</a> | <a href="admin_menpai.asp?id=<%=rs("id")%>&action=del">ɾ��</a> | <%=rs("GroupName")%>
</td>
</tr>
<%
	rs.movenext
	loop
	set rs=nothing
%>
</table>
<%
end sub
sub savegroup()
dim GroupName
GroupName=Checkstr(trim(request("GroupName")))
set rs=conn.execute("select top 1 id from GroupName where   GroupName='"&GroupName&"' order by id desc")
if rs.eof and rs.bof then
	conn.execute("insert into GroupName (GroupName) values ('"&GroupName&"')")
else
	errmsg=errmsg+"<br>"+"<li>���������ͬ��������"
	call dvbbs_error()
	exit sub
end if
set rs=nothing

%>
<center><p><b>��ӳɹ���</b>
<%
end sub
sub savedit()
dim GroupName
GroupName=Checkstr(trim(request("GroupName")))
set rs=conn.execute("select top 1 id from GroupName where   GroupName='"&GroupName&"' order by id desc")
if rs.eof and rs.bof then
		conn.execute("update GroupName set GroupName='"&GroupName&"' where id="&request("id"))
else
	errmsg=errmsg+"<br>"+"<li>�����޸ĳ��Ѵ�����ͬ��������"
	call dvbbs_error()
	exit sub
end if
set rs=nothing
%>
<center><p><b>�޸ĳɹ���</b>
<%
end sub
sub del()
	conn.execute("delete from GroupName where id="&request("id"))
%>
<center><p><b>ɾ���ɹ���</b>
<%
end sub
%>