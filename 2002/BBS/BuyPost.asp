<!-- #include file="conn.asp" -->
<!-- #include file="inc/const.asp" -->
<%
	response.buffer=true
	dim rootid
	dim AnnounceID
	stats="��������"
	if BoardID="" or not isInteger(BoardID) or BoardID=0 then
		Errmsg=Errmsg+"<br>"+"<li>����İ����������ȷ�����Ǵ���Ч�����ӽ��롣"
		founderr=true
	else
		BoardID=clng(BoardID)
	end if
	if request("id")="" then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>��ָ��������ӡ�"
	elseif not isInteger(request("id")) then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>�Ƿ������Ӳ�����"
	else
		rootid=request("id")
	end if
	if request("replyID")="" then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>��ָ��������ӡ�"
	elseif not isInteger(request("replyID")) then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>�Ƿ������Ӳ�����"
	else
		AnnounceID=request("replyID")
	end if
	dim FoundTable
	FoundTable=false
	if request("PostTable")<>"" then
	For i=0 to ubound(AllPostTable)
		if request("PostTable")=AllPostTable(i) then
			FoundTable=true
			exit for
		end if
	next
	else
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>�Ƿ��Ĳ�����"
	end if
	if not FoundTable then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>�Ƿ��Ĳ�����"
	end if
	if not founduser then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>���½����в�����"
	end if

if founderr then
	call nav()
	call head_var(2,0,"","")
	call dvbbs_error()
else
	call nav()
	call head_var(1,boarddepth,"","")
	call main()
	if founderr then call dvbbs_error()
end if
call activeonline()
call footer()

sub main()
	dim re
	dim po,ii
	dim reContent
	dim strContent
	dim PostBuyUser
	po=0
	ii=0
	dim usermoney
	set rs=conn.execute("select userWealth from [user] where userid="&userid)
	usermoney=rs(0)
	set rs=server.createobject("adodb.recordset")
	sql="select body,PostBuyUser,username,PostUserID from "&request("PostTable")&" where Announceid="&Announceid
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>��������Ӳ�����"
		exit sub
	else
	strContent=HTMLcode(rs(0))
	PostBuyUser=rs(1)
	Set re=new RegExp
	re.IgnoreCase =true
	re.Global=True
	re.Pattern="(^.*)(\[UseMoney=*([0-9]*)\])(.*)(\[\/UseMoney\])(.*)"
	po=re.Replace(strContent,"$3")
	if IsNumeric(po) then
		ii=int(po) 
	else
		ii=0
	end if
	set re=nothing
	if membername=rs(2) then
		response.write "<script>alert('�Ǻǣ���Ҫ��Ǯ�����Լ������������');</script>"
	elseif usermoney>ii then
		if (not isnull(PostBuyUser)) and PostBuyUser<>"" then
		if instr("|"&PostBuyUser&"|","|"&membername&"|")>0 then
		response.write "<script>alert('�Ǻǣ����Ѿ��������ѽ��');</script>"
		else
		conn.execute("update [user] set userWealth=userWealth-"&ii&" where userid="&userid)
		conn.execute("update [user] set userWealth=userWealth+"&ii&" where userid="&rs(3))
		rs(1)=rs(1) & "|" & membername
		rs.update
		response.write "<script>alert('����ɹ���');</script>"
		end if
		else
		conn.execute("update [user] set userWealth=userWealth-"&ii&" where userid="&userid)
		conn.execute("update [user] set userWealth=userWealth+"&ii&" where userid="&rs(3))
		rs(1)=membername
		rs.update
		response.write "<script>alert('����ɹ���');</script>"
		end if
	else
		response.write "<script>alert('����û��Ǯѽ��');</script>"
	end if
	end if
	rs.close
	set rs=nothing
	response.redirect "dispbbs.asp?boardid="&request("boardid")&"&id="&rootid&"&replyID="&AnnounceID&"&skin=1"
end sub
%>