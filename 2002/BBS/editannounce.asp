<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/char_board.asp" -->
<%
dim AnnounceID
dim replyID
dim username
dim rs_old
dim old_user
dim con,content
dim topic
dim olduser,oldpass
dim totalusetable
dim CanEditPost
CanEditPost=False
if BoardID="" or not isInteger(BoardID) or BoardID=0 then
	Errmsg=Errmsg+"<br>"+"<li>����İ����������ȷ�����Ǵ���Ч�����ӽ��롣"
	founderr=true
else
	BoardID=clng(BoardID)
end if

if cint(Board_Setting(2))=1 then
	if not founduser then
		Errmsg=Errmsg+"<br>"+"<li>����̳Ϊ��֤��̳����<a href=login.asp>��½</a>��ȷ�������û����Ѿ��õ�����Ա����֤����롣"
		founderr=true
	else
		if chkboardlogin(boardid,membername)=false then
			Errmsg=Errmsg+"<br>"+"<li>����̳Ϊ��֤��̳����ȷ�������û����Ѿ��õ�����Ա����֤����롣"
			founderr=true
		end if
	end if
end if
if request("id")="" then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>��ָ��������ӡ�"
elseif not isInteger(request("id")) then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>�Ƿ������Ӳ�����"
else
	AnnounceID=request("id")
end if
if request("replyID")="" then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>��ָ��������ӡ�"
elseif not isInteger(request("replyID")) then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>�Ƿ������Ӳ�����"
else
	replyID=request("replyID")
end if
if not founduser then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>���½����в�����"
end if
if Cint(Board_Setting(43))=1 then
	Errmsg=Errmsg+"<br>"+"<li>����̳�Ѿ�������Ա�����˲���������"
	founderr=true
end if
if cint(Board_Setting(1))=1 then
	if Cint(GroupSetting(37))=0 then
		Errmsg=ErrMsg+"<Br>"+"<li>��û��Ȩ�޽���������̳��"
		founderr=true
	end if
end if
stats=boardtype & "�༭����"
if founderr then
	call nav()
	call head_var(2,0,"","")
	call dvbbs_error()
else
	call nav()
	call head_var(1,BoardDepth,0,0)
	call main()
	if founderr then call dvbbs_error()
end if
call activeonline()
call footer()

sub main()
set rs=conn.execute("select PostTable from topic where TopicID="&AnnounceID)
if rs.eof and rs.bof then
	Errmsg=Errmsg+"<br>"+"<li>û���ҵ���Ӧ�����ӡ�"
	Founderr=true
	exit sub
else
	TotalUseTable=rs(0)
end if

sql="select b.username,b.topic,b.body,b.dateandtime,u.usergroupID from "&TotalUseTable&" b,[user] u where b.postuserid=u.userid and b.AnnounceID="&replyID
set rs=conn.execute(sql)
if rs.eof and rs.bof then
	Errmsg=Errmsg+"<br>"+"<li>û���ҵ���Ӧ�����ӡ�"
	Founderr=true
	exit sub
else
	topic=rs("topic")
	con=rs("body")
	old_user=rs("username")
	if Clng(forum_setting(51))>0 and not (master or boardmaster or superboardmaster) then
		if Datediff("s",rs("dateandtime"),Now())>Clng(forum_setting(51))*60 then
		Errmsg=Errmsg+"<br>"+"<li>ϵͳ�༭����ʱ��Ϊ"&forum_setting(51)&"���ӣ��������������ӵ������Ѿ���"&Datediff("s",rs("dateandtime"),Now())\60&"�����ˡ�"
			founderr=true
			exit sub
		end if
	end if
	if rs("username")=membername then
		if Cint(GroupSetting(10))=0 then
			Errmsg=Errmsg+"<br>"+"<li>��û���ڱ���̳�༭�Լ����ӵ�Ȩ�ޣ���<a href=login.asp>��½</a>����ͬ����Ա��ϵ��"
			founderr=true
			exit sub
			CanEditPost=False
		else
			CanEditPost=True
		end if
	else
		if (master or superboardmaster or boardmaster) and Cint(GroupSetting(23))=1 then
			CanEditPost=True
		else
			CanEditPost=False
		end if
		if UserGroupID>3 and Cint(GroupSetting(23))=1 then
			CanEditPost=true
		end if
		if Cint(GroupSetting(23))=1 and FoundUserPer then
			CanEditPost=True
		elseif Cint(GroupSetting(23))=0 and FoundUserPer then
			CanEditPost=False
		end if
		if UserGroupID<4 and UserGroupID=rs("UserGroupID") then
			Errmsg=Errmsg+"<br>"+"<li>ͬ�ȼ��û������޸ġ�"
			Founderr=true
			exit sub
		elseif UserGroupID<4 and UserGroupID>rs("UserGroupID") then
			Errmsg=Errmsg+"<br>"+"<li>�����޸ĵȼ������ߵ��û������ӡ�"
			Founderr=true
			exit sub
		end if
		if not CanEditPost then
			Errmsg=Errmsg+"<br>"+"<li>��û���㹻��Ȩ�ޱ༭�����ӣ���͹���Ա��ϵ��"
			Founderr=true
			exit sub
		end if
	end if
end if

set rs=nothing

if con<>"" then
	content=replace(con,"<br>",chr(13))
	content=replace(content,"&nbsp;","")
	content=content+chr(13)
end if
%>
<script src="inc/ubbcode.js"></script>
<form action="SaveditAnnounce.asp?boardID=<%=boardID%>&replyID=<%=replyID%>&ID=<%=announceID%>" method=POST name=frmAnnounce>
  <input type=hidden name=followup value=<%=AnnounceID%>>
  <input type=hidden name="star" value="<%=request("star")%>">
  <input type=hidden name="TotalUseTable" value="<%=TotalUseTable%>">
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
    <tr>
      <th width=100% colspan=2>�༭����</th>
    </tr>
        <tr>
          <td width=20% class=tablebody1><b>�û���</b></td>
          <td width=80% class=tablebody1><input name=username value=<%=htmlencode(old_user)%> >&nbsp;&nbsp;<font color=<%=Forum_body(8)%>><b>*</b></font><a href=reg.asp>��û��ע�᣿</a> 
          </td>
        </tr>
        <tr>
          <td width=20% class=tablebody1><b>����</b></td>
          <td width=80% class=tablebody1><input name=passwd type=password value=<%=htmlencode(memberword)%> >&nbsp;&nbsp;<font color=<%=Forum_body(8)%>><b>*</b></font><a href=lostpass.asp>�������룿</a>�����༭����Ҫ����</td>
        </tr>
        <tr>
          <td width=20% class=tablebody1><b>�������</b>
              <SELECT name=font onchange=DoTitle(this.options[this.selectedIndex].value)>
              <OPTION selected value="">ѡ����</OPTION> <OPTION value=[ԭ��]>[ԭ��]</OPTION> 
              <OPTION value=[ת��]>[ת��]</OPTION> <OPTION value=[��ˮ]>[��ˮ]</OPTION> 
              <OPTION value=[����]>[����]</OPTION> <OPTION value=[����]>[����]</OPTION> 
              <OPTION value=[�Ƽ�]>[�Ƽ�]</OPTION> <OPTION value=[����]>[����]</OPTION> 
              <OPTION value=[ע��]>[ע��]</OPTION> <OPTION value=[��ͼ]>[��ͼ]</OPTION>
              <OPTION value=[����]>[����]</OPTION> <OPTION value=[����]>[����]</OPTION>
              <OPTION value=[����]>[����]</OPTION></SELECT>
         <td width=80% class=tablebody1><input name=subject size=70 maxlength=80 value=<%=htmlencode(Topic)%> >&nbsp;&nbsp;<font color=<%=Forum_body(8)%>><b>*</b></font>���ó��� 25 �����ֻ�50��Ӣ���ַ�
	 </td>
        </tr>
        <tr>
          <td width=20% valign=top class=tablebody1><b>��ǰ����</b><br><li>���������ӵ�ǰ��</td>
          <td width=80% class=tablebody1>
<%for i=0 to Forum_PostFaceNum-1%>
	<input type="radio" value="<%=Forum_PostFace(i)%>" name="Expression" <%if i=0 then response.write "checked"%>><img src="<%=forum_info(8)&Forum_PostFace(i)%>" >&nbsp;&nbsp;
<%if i>0 and ((i+1) mod 9=0) then response.write "<br>"%>
<%next%>
</td>
        </tr>
        <tr>
          <td width=20% valign=top class=tablebody1>
<b>����</b><br>
<li>HTML��ǩ�� <%=iif(Board_Setting(5),"����","������")%>
<li>UBB��ǩ�� <%=iif(Board_Setting(6),"����","������")%>
<li>��ͼ��ǩ�� <%=iif(Board_Setting(7),"����","������")%>
<li>��ý���ǩ��<%=iif(Board_Setting(9),"����","������")%>
<li>�����ַ�ת����<%=iif(Board_Setting(8),"����","������")%>
<li>�ϴ�ͼƬ��<%=iif(Forum_Setting(3),"����","������")%>
<li>���<%=Board_Setting(16)\1024%>KB<BR><BR>
<B>��������</B><BR>
<li><%=iif(Board_Setting(10),"<a href=""javascript:money()"" title=""ʹ���﷨��[money=������ò���������Ҫ��Ǯ��]����[/money]"">��Ǯ��</a>","��Ǯ����������")%>
<li><%=iif(Board_Setting(11),"<a href=""javascript:point()"" title=""ʹ���﷨��[point=������ò���������Ҫ������]����[/point]"">������</a>","��������������")%>
<li><%=iif(Board_Setting(12),"<a href=""javascript:usercp()"" title=""ʹ���﷨��[usercp=������ò���������Ҫ������]����[/usercp]"">������</a>","��������������")%>
<li><%=iif(Board_Setting(13),"<a href=""javascript:power()"" title=""ʹ���﷨��[power=������ò���������Ҫ������]����[/power]"">������</a>","��������������")%>
<li><%=iif(Board_Setting(14),"<a href=""javascript:article()"" title=""ʹ���﷨��[post=������ò���������Ҫ������]����[/post]"">������</a>","��������������")%>
<li><%=iif(Board_Setting(15),"<a href=""javascript:replyview()"" title=""ʹ���﷨��[replyview]�ò������ݻظ���ɼ�[/replyview]"">�ظ���</a>","�ظ�����������")%>
<li><%=iif(Board_Setting(23),"<a href=""javascript:usemoney()"" title=""ʹ���﷨��[usemoney=����ò���������Ҫ���ĵĽ�Ǯ��]����[/usemoney]"">������</a>","��������������")%>
	  </td>
          <td width=80% class=tablebody1>
<%if Cint(Board_Setting(4))=1 then%>
<!--#include file="inc/getubb.asp"-->
<%end if%>
<textarea class=smallarea cols=95 name=Content rows=12 wrap=VIRTUAL  onkeydown=ctlent()>
<%=server.htmlencode(content)%>
</textarea>
	</td>
        </tr>
		<tr>
<td class=tablebody1 valign=top colspan=2 style="table-layout:fixed; word-break:break-all"><b>�������ͼ�����������м�����Ӧ�ı���</B><br>&nbsp;
<%
dim ii
for i=0 to 12
	if len(i)=1 then ii="0" & i else ii=i
	response.write "<img src="""&Forum_info(10)&Forum_emot(i)&""" border=0 onclick=""insertsmilie('[em"&ii&"]')"" style=""CURSOR: hand"">&nbsp;"
next
%>
    		</td>
                </tr>
                <tr>
                <td valign=top class=tablebody1><b>ѡ��</b></td>
                <td valign=middle class=tablebody1><input type=checkbox name=signflag value=yes checked>�Ƿ���ʾ����ǩ����<br>
                <input type=checkbox name=emailbody value=yes>�лظ�ʱʹ���ʼ�֪ͨ����
<BR><BR></td>
                </tr><tr>
                <td class=tablebody2 valign=middle colspan=2 align=center>
                <input type=Submit value="�� ��" name=Submit> &nbsp; <input type=button value='Ԥ ��' name=Button onclick=gopreview()>&nbsp;<input type=reset name=Submit2 value="�� ��">
                </td></form></tr>
      </table>
</form>
<form name=preview action=preview.asp?boardid=<%=boardid%> method=post target=preview_page>
<input type=hidden name=title value=><input type=hidden name=body value=>
</form>

<%
end sub
%>