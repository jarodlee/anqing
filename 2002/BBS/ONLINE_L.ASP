<%
Rem ==========论坛首页函数=========
Rem 判断用户系统信息
function usersysinfo(info,getinfo)
if instr(info,";")>0 then
	dim usersys
	usersys=split(info,";")
	if ubound(usersys)>=2 then
	usersys(1)=replace(usersys(1),"MSIE","Internet Explorer")
	usersys(2)=replace(usersys(2),")","")
	usersys(2)=replace(usersys(2),"NT 5.1","XP")
	usersys(2)=replace(usersys(2),"NT 5.0","2000")
	usersys(2)=replace(usersys(2),"9x","Me")
	usersys(1)="浏 览 器：" & Trim(usersys(1))
	usersys(2)="操作系统：" & Trim(usersys(2))
	if getinfo=1 then
		usersysinfo=usersys(1)
	else
		usersysinfo=usersys(2)
	end if
	else
	if getinfo=1 then
		usersysinfo="浏 览 器：未知"
	else
		usersysinfo="操作系统：未知"
	end if
	end if
else
	if getinfo=1 then
		usersysinfo="未知"
	else
		usersysinfo="未知"
	end if
end if
end function

function online(boardid)
dim tmprs
if boardid=0 then
	sql="Select count(*) from online where userid>0"
else
	sql="Select count(*) from online where userid>0 and boardid="&boardid
end if
set tmprs=conn.execute(sql) 
online=tmprs(0) 
set tmprs=nothing 
if isnull(online) then online=0
end function 

function guest(boardid)
dim tmprs
if boardid=0 then
	sql="Select count(*) from online where userid=0"
else
	sql="Select count(*) from online where userid=0 and boardid="&boardid
end if
set tmprs=conn.execute(sql) 
guest=tmprs(0) 
set tmprs=nothing
if isnull(guest) then guest=0
end function

sub onlineuser(online_u,online_g,boardid)
if boardid>0 and (cint(online_u)=1 or cint(online_g)=1) then
response.write "<tr><td colspan=2 class=tablebody1><table cellpadding=2 cellspacing=1 border=0 width=""100%"" style=""word-break:break-all;""><tr>"
end if
dim online_face
dim sip,acip,userip
dim NowStats,ActiveTime,Binfo,UComeFrom,BrStr
if cint(online_u)=1 then
	i=0
	'用户信息
	if boardid=0 then
	sql="select username,ip,stats,UserGroupID,userhidden,userid,startime,lastimebk,actforip,browser from online where userid>0"
	else
	sql="select username,ip,stats,UserGroupID,userhidden,userid,startime,lastimebk,actforip,browser from online where boardid="&boardid&" and userid>0"
	end if
	set rs=conn.execute(sql)
	do while not rs.eof
	sip=rs(1)
	acip=rs(8)
	if membername=rs(0) then
		Brstr="&#13;&#10;"
	else
		Brstr="<br>"
	end if
	if Cint(forum_setting(33))=1 then
		NowStats="目前位置：" & htmlencode(rs(2)) & Brstr
	else
		NowStats=""
	end if
	if Cint(forum_setting(34))=1 then
		ActiveTime="来访时间：" & rs(6) & Brstr & "活动时间：" & rs(7) & Brstr
	else
		ActiveTime=""
	end if
	if Cint(forum_setting(35))=1 then
		Binfo=usersysinfo(rs(9),2) & Brstr & usersysinfo(rs(9),1) & Brstr
	else
		Binfo=""
	end if
	if Cint(GroupSetting(30))=0 then
		userip="真实ＩＰ：已设置保密"
	else
		if acip <> "" then
			userip="真实ＩＰ：" & acip
		else
			userip="真实ＩＰ：" & sip
		end if
	end if
	if cint(forum_setting(36))=1 then
		if Cint(GroupSetting(30))=0 then
			UComeFrom=Brstr & "用户来源：已设置保密"
		else
			if acip<>"" then
				UComeFrom=Brstr & "用户来源：" & address(acip)
			else
				UComeFrom=Brstr & "用户来源：" & address(sip)
			end if
		end if
	else
		UComeFrom=""
	end if
	select case rs(3)
	case 1
	online_face="<img src="&Forum_info(7)&Forum_pic(0)&" alt="&rs(5)&" width=12 height=11>"
	case 2
	online_face="<img src="&Forum_info(7)&Forum_pic(1)&" alt="&rs(5)&" width=12 height=11>"
	case 3
	online_face="<img src="&Forum_info(7)&Forum_pic(1)&" alt="&rs(5)&" width=12 height=11>"
	case 8
	online_face="<img src="&Forum_info(7)&Forum_pic(2)&" alt="&rs(5)&" width=12 height=11>"
	case else
	online_face="<img src="&Forum_info(7)&Forum_pic(3)&" width=12 height=11>&nbsp;"
	end select

	if membername=rs(0) then
		response.write "<td width=""14%"">" & online_face&"&nbsp;<a href=dispuser.asp?id="&rs(5)&" target=_blank title="""& NowStats & ActiveTime & Binfo & UserIP & UComeFrom &"""><font color="&Forum_pic(5)&">"&htmlencode(rs(0))&"</font></a></td>"
	else
		if rs(4)=1 then
			if boardmaster or master then
			response.write "<td width=""14%"">" & online_face&"&nbsp;<a href=dispuser.asp?id="&rs(5)&" target=_blank title="""& NowStats & ActiveTime & Binfo & UserIP & UComeFrom &""">"&htmlencode(rs(0))&"</a></td>"
			else
			response.write "<td width=""14%""><img src="&Forum_info(7)&Forum_pic(4)&" width=12 height=11>&nbsp;<a href=# target=_blank title="""& NowStats & ActiveTime & Binfo & UserIP & UComeFrom &""">隐身会员</a></td>"
			end if
		else
			response.write "<td width=""14%"">" & online_face&"&nbsp;<a href=dispuser.asp?id="&rs(5)&" target=_blank title="""& NowStats & ActiveTime & Binfo & UserIP & UComeFrom &""">"&htmlencode(rs(0))&"</a></td>"
		end if
	end if
	if i=6 then response.write "</tr><tr>"
	if i>6 then 
		i=1
	else
		i=i+1
	end if
	rs.movenext
	loop
end if
if cint(online_g)=1 then
	dim onlineusername
	i=0
	if boardid=0 then
	sql="select username,ip,stats,UserGroupID,userhidden,userid,startime,lastimebk,actforip,id,browser from online where userid=0"
	else
	sql="select username,ip,stats,UserGroupID,userhidden,userid,startime,lastimebk,actforip,id,browser from online where boardid="&boardid&" and userid=0"
	end if
	set rs=conn.execute(sql)
  
	online_face="<img src="&Forum_info(7)&Forum_pic(4)&" width=12 height=11>"
	if not (rs.eof and rs.eof) then
		response.write "</tr><tr>"
	end if
	do while not rs.eof
	sip=rs(1)
	acip=rs(8)
	if trim(session("userid"))<>"" and isnumeric(session("userid")) then
		if int(session("userid"))=int(rs(9)) then
			Brstr="&#13;&#10;"
			onlineusername="<font color="&Forum_pic(5)&">客人</font>"
		else
			Brstr="<br>"
			onlineusername="客人"
		end if
	else
		Brstr="<br>"
		onlineusername="客人"
	end if
	if Cint(forum_setting(33))=1 then
		NowStats="目前位置：" & htmlencode(rs(2)) & Brstr
	else
		NowStats=""
	end if
	if Cint(forum_setting(34))=1 then
		ActiveTime="来访时间：" & rs(6) & Brstr & "活动时间：" & rs(7) & Brstr
	else
		ActiveTime=""
	end if
	if Cint(forum_setting(35))=1 then
		Binfo=usersysinfo(rs(10),2) & Brstr & usersysinfo(rs(10),1) & Brstr
	else
		Binfo=""
	end if
	if Cint(GroupSetting(30))=0 then
		userip="真实ＩＰ：已设置保密"
	else
		if acip <> "" then
			userip="真实ＩＰ：" & acip
		else
			userip="真实ＩＰ：" & sip
		end if
	end if
	if cint(forum_setting(36))=1 then
		if Cint(GroupSetting(30))=0 then
			UComeFrom=Brstr & "用户来源：已设置保密"
		else
			if acip<>"" then
				UComeFrom=Brstr & "用户来源：" & address(acip)
			else
				UComeFrom=Brstr & "用户来源：" & address(sip)
			end if
		end if
	else
		UComeFrom=""
	end if
	response.write "<td width=""14%"">" & online_face&"&nbsp;<a href=# target=_blank title="""& NowStats & ActiveTime & Binfo & UserIP & UComeFrom &""">"&onlineusername&"</a></td>"

	if i=6 then response.write "</tr><tr>"
	if i>6 then 
		i=1
	else
		i=i+1
	end if
	rs.movenext
	loop
end if
if boardid>0 and (cint(online_u)=1 or cint(online_g)=1) then
	response.write "</tr></TABLE>"
end if
set rs=nothing
end sub
%>