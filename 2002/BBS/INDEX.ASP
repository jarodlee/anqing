<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
response.buffer=true
'=========================================================
' File: index.asp
' Version:5.0
' Date: 2002-9-21
' Script Written by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================

Dim OnlineNum,GuestNum
Dim reBoard_Setting,BoardCount,k,ColSpanNum
Dim Board_Count
Dim TdTableWidth
OnlineNum=online(0)
GuestNum=guest(0)
Stats="论坛首页"
call activeonline()
Rem 首页顶部信息
call nav()
if isnull(lastlogin) or lastlogin="" then lastlogin=now()
if Clng(Forum_OnlineNum)>Clng(Maxonline) then
	conn.execute("update config set Maxonline="&allonline()&",MaxonlineDate=Now()")
end if
%>
<TABLE cellSpacing=0 cellPadding=0 width=<%=Forum_body(12)%> border=0 align=center>
<%if founduser then%>
<tr><td width=65% >
</td><td width=35% align=right>
<%if Cint(newincept())>Cint(0) then%>
<%if Cint(forum_setting(10))=1 then%>
<script language=JavaScript>openScript('messanger.asp?action=read&id=<%=inceptid(1)%>&sender=<%=inceptid(2)%>',500,400)</script>
<%end if%>
<bgsound src="<%=Forum_info(7)&Forum_statePic(8)%>" border=0>
<img src=<%=Forum_info(7)&Forum_boardpic(9)%>> <a href="usersms.asp?action=inbox">我的收件箱</a> (<a href="javascript:openScript('messanger.asp?action=read&id=<%=inceptid(1)%>&sender=<%=inceptid(2)%>',500,400)"><font color="<%=Forum_body(8)%>"><%=newincept()%> 新</font></a>)
<%else%>
<img src=<%=Forum_info(7)&Forum_boardpic(8)%>> <a href="usersms.asp?action=inbox">我的收件箱</a> (<font color=gray>0 新</font>)
<%end if%>
<%else%>
<br>
<%end if%>
</td></tr>
<tr>
<td align=center width=100% valign=middle colspan=2>
<SCRIPT LANGUAGE='JavaScript' SRC='js/fader.js' TYPE='text/javascript'></script>
<SCRIPT LANGUAGE='JavaScript' TYPE='text/javascript'>
prefix="";
arNews = ["<%
myCache.name="AnnounceMents"&BoardID
if myCache.valid then
	response.write myCache.value
else
Dim AnnounceMent_list
sql="select top 1 title,addtime from bbsnews where boardid=0 order by id desc"
set rs=conn.execute(sql)
if rs.bof and rs.eof then
	AnnounceMent_list="<b><a href='announcements.asp?boardid="& boardid &"' target=_blank><ACRONYM TITLE='当前没有公告'>当前没有公告</ACRONYM></a></b> ("& now() &")"
else
	AnnounceMent_list="<b><a href='announcements.asp?boardid="& boardid &"' target=_blank>"& rs("title") &"</a></b> ("& rs("addtime") &")"
end if
myCache.add AnnounceMent_list,dateadd("n",9999,now)
response.write AnnounceMent_list
set rs=nothing
end if
%>","","欢迎光临<%=Forum_info(0)%> <%=boardtype%>","",
"严禁恶意使用粗言秽语，违者经劝告无效，立即封ID。",""
]
</SCRIPT>
<div id="elFader" style="position:relative;visibility:hidden; height:16" ></div>
</td>
</tr>
</table>
<TABLE border=0 width="<%=Forum_body(12)%>" align=center>
<tr><td colspan=2 height=1 bgcolor="<%=Forum_body(27)%>">

</td></tr>
<tr><td colspan=2 height=5>

</td></tr>
<%
sql="select top 1 TopicNum,BbsNum,TodayNum,UserNum,lastUser,yesterdaynum,maxpostnum,maxpostdate,LastPost from config where active=1"
set rs=conn.execute(sql)
LastPostInfo=split(rs("LastPost"),"$")
if not isdate(LastPostInfo(2)) then LastPostInfo(2)=Now()
LastPostTime=LastPostInfo(2)
if datediff("d",LastPostTime,Now())>0 and ubound(LastPostInfo)=7 then
	conn.execute("update config set yesterdaynum=todaynum,LastPost='"&checkStr(LastPostInfo(0))&"$"&checkStr(LastPostInfo(1))&"$"&Now()&"$"&checkStr(LastPostInfo(3))&"$"&checkStr(LastPostInfo(4))&"$"&checkStr(LastPostInfo(5))&"$"&checkStr(LastPostInfo(6))&"$"&checkStr(LastPostInfo(7))&"'")
	conn.execute("update config set todaynum=0")
end if
if rs(2)>rs(6) then conn.execute("update config set maxpostnum=todaynum,maxpostdate=Now()")
%>
<TR><TD style="line-height: 20px;">
欢迎新加入会员 <a href=dispuser.asp?name=<%= htmlencode(rs(4)) %> target=_blank><b><%= htmlencode(rs(4)) %></b></a>&nbsp;[<a href="toplist.asp?orders=2">新进来宾</a>]<BR>论坛共有 <B><%= rs(3)%></B> 位注册会员 , 主题总数：<b><%= rs(0) %></b> , 帖子总数：<b><%= rs(1) %></b><BR>今日论坛共发贴：<FONT COLOR="<%=Forum_body(8)%>"><B><%=rs(2)%></B></FONT> , 昨日发贴：<B><%=rs(5)%></B> , 最高日发贴：<B><%=rs(6)%></B></td><TD valign=bottom align=right style="line-height: 20px;"><a href=queryresult.asp?stype=3>查看新贴</a> , <a href=hotlist.asp?stype=1>热门话题</a> , <a href=toplist.asp?orders=1>发贴排行</a> , <a href=toplist.asp?orders=7>用户列表</a><BR>您最后一次访问是在：<%=Now()%>
</TD></TR>
</TABLE>
<%
if Cint(Forum_setting(4))=1 then
call main_2()
else
call main_1()
end if

sub main_1()
if not founduser and Cint(forum_setting(31))=1 then
	call FastLogin()
end if
Dim cBoardID
Response.Cookies("BoardList").expires= date+7
cBoardID= Request("cBoardid")
if isnumeric(cBoardID) then
	if request("CatLog")="N" then
	Response.Cookies("BoardList")(cBoardID & "BoardID")= "NotShow"
	else
	Response.Cookies("BoardList")(cBoardID & "BoardID")= "Show"
	end if
end if
dim simdisp
simdisp=false
k=0
'只列出深度小于2的版面
sql="select * from board where depth<="&forum_setting(5)&" order by rootid,orders"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if not (rs.eof and rs.bof) then BoardCount=rs.recordcount else boardcount=0 end if
do while not rs.eof
k=k+1
reBoard_Setting=split(rs("Board_setting"),",")
if rs("BoardMaster")<>"" and not isnull(rs("BoardMaster")) then
	master_1=split(rs("BoardMaster"), "|")
	for i = 0 to ubound(master_1)
	if i>6 then
		master_2=master_2
	else
		master_2=master_2 & "<a href=dispuser.asp?name="&master_1(i)&" target=_blank>"&master_1(i)&"</a>&nbsp;"
	end if
	next
	if i>7 then master_2=master_2 & "<font color=gray>More...</font>"
else
	master_2="暂无"
end if
if rs("ParentID")=0 then
	if request.cookies("BoardList")(rs("boardid") & "BoardID")="NotShow" then
		ColSpanNum=reBoard_Setting(41)
		if not isnumeric(ColSpanNum) and Cint(ColSpanNum)=0 then ColSpanNum=4
		simdisp=true
	elseif request.cookies("BoardList")(rs("boardid") & "BoardID")="Show" then
		simdisp=false
		ColSpanNum=2
	else
		if Cint(reBoard_Setting(39))=1 then
			ColSpanNum=reBoard_Setting(41)
			if not isnumeric(ColSpanNum) and Cint(ColSpanNum)=0 then ColSpanNum=4
			simdisp=true
		else
			simdisp=false
			ColSpanNum=2
		end if
	end if
	response.write "<table cellspacing=1 cellpadding=0 align=center class=tableBorder1>"
	response.write "<TR><Th colSpan="&ColSpanNum&" height=25 align=left id=TableTitleLink>&nbsp;"
	if simdisp then
		response.write "<a href=""?cBoardid="&rs("boardid")&"&Catlog=Y"" title=""展开论坛列表""><img src="""&Forum_info(7)&Forum_boardpic(15)&""" border=0></a><a href=list.asp?boardid="&rs("boardid")&" title=进入本分类论坛>"&rs("boardtype")&"</a>"
	else
		response.write "<a href=""?cBoardid="&rs("boardid")&"&Catlog=N"" title=""关闭论坛列表""><img src="""&Forum_info(7)&Forum_boardpic(16)&""" border=0></a><a href=list.asp?boardid="&rs("boardid")&" title=进入本分类论坛>"&rs("boardtype")&"</a>"
	end if
	response.write "</Th></TR>"
	Board_Count=0
else
	if simdisp then
		call main_1_boardlist_2()
	else
		call main_1_boardlist_1()
	end if
end if
master_2=""
rs.movenext
if k<BoardCount then
	if rs("parentid")=0 and Board_count<ColSpanNum and simdisp then
		for i=Board_Count+1 to ColSpanNum
			response.write "<td class=tablebody1 width="""&TdTableWidth&"%"">&nbsp;</td>"
		next
		response.write "</tr>"
	end if
	if rs("parentid")=0 then
		response.write "</table><BR>"
		simdisp=false
		ColSpanNum=2
	end if
else
	if Board_count<ColSpanNum and simdisp then
		for i=Board_Count+1 to ColSpanNum
			response.write "<td class=tablebody1 width="""&TdTableWidth&"%"">&nbsp;</td>"
		next
		response.write "</tr>"
	end if
end if
loop
set rs=nothing
response.write "</table>"
if not founduser and Cint(forum_setting(31))=2 then
	call FastLogin()
end if
end sub

sub main_1_boardlist_1()
if (Cint(reBoard_Setting(1))=1 and Cint(GroupSetting(37))=1) or Cint(reBoard_Setting(1))=0 then
response.write "<TR><TD align=middle width=100% class=tablebody1><table width=""100%"" cellspacing=0 cellpadding=0>"
LastPostInfo=split(rs("LastPost"),"$")
if not isdate(LastPostInfo(2)) then LastPostInfo(2)=Now()
'=========================BoardInfo============================
response.write "<TR><TD align=middle width=46 class=tablebody1>"
if Cint(reBoard_Setting(0))=1 then
	if datediff("h",lastlogin,LastPostInfo(2))=0 then
	response.write "<img src="&Forum_info(7)&Forum_pic(15)&" alt=有新帖子>"
	else
	response.write "<img src="&Forum_info(7)&Forum_pic(14)&" alt=无新帖子>"
	end if
else
	if datediff("h",lastlogin,LastPostInfo(2))=0 then
	response.write "<img src="&Forum_info(7)&Forum_pic(7)&" alt=有新帖子>"
	else
	response.write "<img src="&Forum_info(7)&Forum_pic(6)&" alt=无新帖子>"
	end if
end if
%>
</TD>
<TD width=1 bgcolor=<%=Forum_body(27)%>>
<TD vAlign=top width=* class=tablebody1>
<TABLE cellSpacing=0 cellPadding=2 width=100% border=0>
<tr><td class=tablebody1 width=*><a href="list.asp?boardid=<%=rs(0)%>">
<font color=<%=Forum_body(14)%>><%=rs("boardtype")%></font></a>
<%
if rs("child")>0 then
response.write "(<font title=""有"&rs("child")&"个下属论坛"">"&rs("child")&")</font>"
end if
%>
</td>
<%
response.write "<td width=40 rowspan=2 align=center class=tablebody1>"
if rs("indeximg")<>"" then
	response.write "<table align=left><tr><td><a href=list.asp?boardid="&rs("boardid")&"><img src="&rs("indeximg")&" align=top border=0></a></td><td width=20></td></tr></table>"
end if
response.write "</td><td width=200 rowspan=2 class=tablebody1>"
if Cint(reBoard_Setting(2))=1 then
response.write "认证论坛，请认证用户进入浏览"
else
response.write "主题：<a href=""dispbbs.asp?Boardid="&LastPostInfo(7)&"&ID="&LastPostInfo(6)&"&replyID="&LastPostInfo(1)&"&skin=1"">"&htmlencode(left(LastPostInfo(3),10))&"..</a><BR>作者：<a href=""dispuser.asp?id="&htmlencode(LastPostInfo(5))&""" target=_blank>"&htmlencode(LastPostInfo(0))&"</a><BR>日期：" & FormatDateTime(LastPostInfo(2),1) & "" & FormatDateTime(LastPostInfo(2),4)&"&nbsp;<a href=""dispbbs.asp?Boardid="&LastPostInfo(7)&"&ID="&LastPostInfo(6)&"&replyID="&LastPostInfo(1)&"&skin=1""><IMG border=0 src="""&Forum_info(7)&Forum_statePic(5)&""" title=""转到："&htmlencode(LastPostInfo(3))&"""></a>"
end if
response.write "</TD></TR><TR><TD width=*><FONT face=Arial><img src="&Forum_info(7)&Forum_pic(8)&" align=middle> "
response.write rs("readme")
response.write"</FONT></TD>"

response.write "</TR><TR><TD class=tablebody2 height=20 width=*>版主："&master_2&""

response.write "</TD><td width=40 align=center class=tablebody2>&nbsp;</td><TD vAlign=middle class=tablebody2 width=200><table width=100% border=0><tr>"

response.write "<td width=25% vAlign=middle><img src="&Forum_info(7)&Forum_pic(9)&" alt=今日帖 align=absmiddle>&nbsp;<font color="&Forum_body(8)&">"&rs("TodayNum")&"</font></td><td width=30% vAlign=middle><img src="&Forum_info(7)&Forum_pic(11)&" alt=主题 border=0  align=absmiddle>&nbsp;"&rs("LastTopicNum")&"</td><td width=45% vAlign=middle><img src="&Forum_info(7)&Forum_pic(10)&" alt=文章 border=0 align=absmiddle>&nbsp;"&rs("LastBBSNum")&"</td></tr>"
response.write "</table></TD></TR></TBODY></TABLE></td></tr></table></td></tr>"
'============================End===============================
end if
end sub

sub main_1_boardlist_2()
TdTableWidth=100/ColSpanNum
if Board_Count=ColSpanNum or Board_Count=0 then response.write "<tr>"
response.write "<td class=tablebody1 width="""&TdTableWidth&"%"">"
	if (Cint(reBoard_Setting(1))=1 and Cint(GroupSetting(37))=1) or Cint(reBoard_Setting(1))=0 then
	response.write "<TABLE cellSpacing=2 cellPadding=2 width=100% border=0><tr><td width=""100%"" title="""&htmlencode(rs("readme"))&"<br>版主："&master_2&""" colspan=2><a href=list.asp?boardid="&rs("boardid")&"><font color="&Forum_body(14)&">"&rs("boardtype")&"</font></a>"
	if rs("child")>0 then
	response.write "(<font title=""有"&rs("child")&"个下属论坛"">"&rs("child")&"</font>)"
	end if
	response.write "</td></tr><tr><td width=""50%"">今日：<font color="&Forum_body(8)&">"&rs("TodayNum")&"</font></td><td width=""50%"">发贴："&rs("LastBbsNum")&"</td></tr></table>"
	end if
response.write "</td>"
if Board_Count=ColSpanNum-1 then response.write "</tr>"
if Cint(Board_Count)>Cint(ColSpanNum)-1 then 
	Board_Count=1
else
	Board_Count=Board_Count+1
end if
end sub

sub main_2()
if not founduser and Cint(forum_setting(31))=1 then
	call FastLogin()
end if
%>
<table cellspacing=1 cellpadding=3 align=center class=tableBorder1>
<TR>
<Th  width=30 height=25>状态</Th>
<Th vAlign=center width=*>论坛名称</Th>
<Th vAlign=center align=middle width=80>版主</Th>
<Th vAlign=center noWrap align=middle width=38>主题 </Th>
<Th vAlign=center noWrap align=middle width=38>贴子 </Th>
<Th vAlign=center noWrap align=middle width=168>最后发表 </Th>
</TR>
<%
Dim cBoardID
Response.Cookies("BoardList").expires= date+7
cBoardID= Request("cBoardid")
if isnumeric(cBoardID) then
	if request("CatLog")="N" then
	Response.Cookies("BoardList")(cBoardID & "BoardID")= "NotShow"
	else
	Response.Cookies("BoardList")(cBoardID & "BoardID")= "Show"
	end if
end if
dim simdisp
simdisp=false
k=0

'只列出深度小于2的版面
sql="select * from board where depth<="&forum_setting(5)&" order by rootid,orders"
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if not (rs.eof and rs.bof) then BoardCount=rs.recordcount else boardcount=0 end if
do while not rs.eof
k=k+1
reBoard_Setting=split(rs("Board_setting"),",")
if rs("BoardMaster")<>"" and not isnull(rs("BoardMaster")) then
	master_1=split(rs("BoardMaster"), "|")
	for i = 0 to ubound(master_1)
	if i>6 then
		master_2=master_2
	else
		master_2=master_2 & "<a href=dispuser.asp?name="&master_1(i)&" target=_blank>"&master_1(i)&"</a><br>"
	end if
	next
	if i>2 then master_2=master_2 & "<font color=gray>More...</font>"
else
	master_2="暂无"
end if

if rs("ParentID")=0 then
	if request.cookies("BoardList")(rs("boardid") & "BoardID")="NotShow" then
		ColSpanNum=reBoard_Setting(41)
		if not isnumeric(ColSpanNum) and Cint(ColSpanNum)=0 then ColSpanNum=4
		simdisp=true
	elseif request.cookies("BoardList")(rs("boardid") & "BoardID")="Show" then
		simdisp=false
		ColSpanNum=6
	else
		if Cint(reBoard_Setting(39))=1 then
			ColSpanNum=reBoard_Setting(41)
			if not isnumeric(ColSpanNum) and Cint(ColSpanNum)=0 then ColSpanNum=4
			simdisp=true
		else
			simdisp=false
			ColSpanNum=6
		end if
	end if
	response.write "<TR><Th colSpan=6 height=25 align=left id=TableTitleLink>&nbsp;"
	if simdisp then
		response.write "<a href=""?cBoardid="&rs("boardid")&"&Catlog=Y"" title=""展开论坛列表""><img src="""&Forum_info(7)&Forum_boardpic(15)&""" border=0></a><a href=list.asp?boardid="&rs("boardid")&" title=进入本分类论坛>"&rs("boardtype")&"</a>"
	else
		response.write "<a href=""?cBoardid="&rs("boardid")&"&Catlog=N"" title=""关闭论坛列表""><img src="""&Forum_info(7)&Forum_boardpic(16)&""" border=0></a><a href=list.asp?boardid="&rs("boardid")&" title=进入本分类论坛>"&rs("boardtype")&"</a>"
	end if
	response.write "</Th></TR>"
	Board_Count=0
else
	if simdisp then
		call main_2_boardlist_2()
	else
		call main_2_boardlist_1()
	end if
end if
master_2=""
rs.movenext
if k<BoardCount then
	if rs("parentid")=0 and Board_count<ColSpanNum and simdisp then
		for i=Board_Count+1 to ColSpanNum
			response.write "<td class=tablebody1 width="""&TdTableWidth&"%"">&nbsp;</td>"
		next
		response.write "</tr></table></td></tr>"
	end if
	if rs("parentid")=0 then
		simdisp=false
		ColSpanNum=6
	end if
else
	if Board_count<ColSpanNum and simdisp then
		for i=Board_Count+1 to ColSpanNum
			response.write "<td class=tablebody1 width="""&TdTableWidth&"%"">&nbsp;</td>"
		next
		response.write "</tr></table></td></tr>"
	end if
end if
loop
set rs=nothing
if not founduser and Cint(forum_setting(31))=2 then
	call FastLogin()
end if
response.write "</table>"
set rs=nothing
end sub

sub main_2_boardlist_1()
dim boardstat
if (Cint(reBoard_Setting(1))=1 and Cint(GroupSetting(37))=1) or Cint(reBoard_Setting(1))=0 then
LastPostInfo=split(rs("LastPost"),"$")
if not isdate(LastPostInfo(2)) then LastPostInfo(2)=Now()
'=========================BoardInfo============================
response.write "<TR><TD width=26 class=tablebody2 valign=middle align=center>"
if Cint(reBoard_Setting(0))=1 then
	if datediff("h",lastlogin,LastPostInfo(2))=0 then
	response.write "<img src="&Forum_info(7)&Forum_pic(15)&" alt=有新帖子>"
	else
	response.write "<img src="&Forum_info(7)&Forum_pic(14)&" alt=无新帖子>"
	end if
else
	if datediff("h",lastlogin,LastPostInfo(2))=0 then
	response.write "<img src="&Forum_info(7)&Forum_pic(7)&" alt=有新帖子>"
	else
	response.write "<img src="&Forum_info(7)&Forum_pic(6)&" alt=无新帖子>"
	end if
end if
%>
</TD><TD vAlign=top width=* class=tablebody1>
<table width=100% cellpadding=2 cellspacing=0>
<tr><td class=tablebody1><a href="list.asp?boardid=<%=rs("boardid")%>">
<font color=<%=Forum_body(14)%>><%=rs("boardtype")%></font></a>
</td></tr>
<tr><td class=tablebody1>
<%
if rs("indeximg")<>"" then
	response.write "<table align=left><tr><td><a href=list.asp?boardid="&rs("boardid")&"><img src="&rs("indeximg")&" align=top border=0></a></td><td width=20></td></tr></table>"
end if
%>
<%=rs("readme")%></td></tr></table>
</TD><TD vAlign=center align=middle class=tablebody2 width=80>
<%=master_2%>
</TD>
<TD vAlign=center noWrap align=middle width=38 class=tablebody1><%=rs("LastTopicNum")%></TD>
<TD vAlign=center noWrap align=middle width=38 class=tablebody1><%=rs("LastBbsNum")%></TD>
<TD noWrap width=168 class=tablebody2>
作者：<a href="dispuser.asp?id=<%=htmlencode(LastPostInfo(5))%>" target=_blank><%=htmlencode(LastPostInfo(0))%></a><a href="dispbbs.asp?Boardid=<%=LastPostInfo(7)%>&ID=<%=LastPostInfo(6)%>&replyID=<%=LastPostInfo(1)%>&skin=1"><IMG border=0 src="<%=Forum_info(7)&Forum_statePic(5)%>" title="转到：<%=htmlencode(LastPostInfo(3))%>"></a><br>
主题：<a href="dispbbs.asp?Boardid=<%=LastPostInfo(7)%>&ID=<%=LastPostInfo(6)%>&replyID=<%=LastPostInfo(1)%>&skin=1"><%=htmlencode(left(LastPostInfo(3),10))%>..</a><br>
<%= FormatDateTime(LastPostInfo(2),1) %><%= FormatDateTime(LastPostInfo(2),4) %>
</TD></TR>
<%
end if
end sub

sub main_2_boardlist_2()
TdTableWidth=100/ColSpanNum
if Board_Count=ColSpanNum or Board_Count=0 then response.write "<tr><td colspan=6 class=tablebody1 width=""100%""><table width=""100%""><tr>"
response.write "<td class=tablebody1 width="""&TdTableWidth&"%"">"
	if (Cint(reBoard_Setting(1))=1 and Cint(GroupSetting(37))=1) or Cint(reBoard_Setting(1))=0 then
	response.write "<TABLE cellSpacing=2 cellPadding=2 width=100% border=0><tr><td width=""100%"" title="""&htmlencode(rs("readme"))&"<br>版主："&master_2&""" colspan=2><a href=list.asp?boardid="&rs("boardid")&"><font color="&Forum_body(14)&">"&rs("boardtype")&"</font></a>"
	if rs("child")>0 then
	response.write "(<font title=""有"&rs("child")&"个下属论坛"">"&rs("child")&"</font>)"
	end if
	response.write "</td></tr><tr><td width=""50%"">今日：<font color="&Forum_body(8)&">"&rs("TodayNum")&"</font></td><td width=""50%"">发贴："&rs("LastBbsNum")&"</td></tr></table>"
	end if
response.write "</td>"
if Board_Count=ColSpanNum-1 then response.write "</tr>"
if Cint(Board_Count)>Cint(ColSpanNum)-1 then 
	Board_Count=1
else
	Board_Count=Board_Count+1
end if
end sub

response.write "<BR><table cellpadding=3 cellspacing=1 align=center class=tableborder1>"&_
	"<TR><Th  colSpan=2 align=left height=25>-=> 友情论坛</Th></TR>"&_
	"<TR><TD vAlign=top class=tablebody1 width=100% >"&_
	"<table width=100% >"

	myCache.name="bbslink"
	if myCache.valid then
	response.write myCache.value
	'mycache.makeEmpty
	else
	Dim linkinfo
	i=7
	sql="select boardname,readme,url from bbslink where islogo=0 order by id"
	set rs=conn.execute(sql)
	if not rs.eof and not rs.bof then
	do while not rs.eof
	if i>6 then 
	linkinfo = linkinfo & "<tr><td width=""16%"">"
	i=1
	else
	linkinfo = linkinfo & "<td width=""16%"">"
	end if
	linkinfo = linkinfo & "<a href="&rs(2)&" target=_blank title="""&rs(1)&""">"&rs(0)&"</a>"
	rs.movenext
	linkinfo = linkinfo & "</td>"
	if i=6 then linkinfo = linkinfo & "</tr>"
	i=i+1
	loop
	end if
	i=7
	sql="select boardname,readme,url,logo from bbslink where islogo=1 order by id"
	set rs=conn.execute(sql)
	if not rs.eof and not rs.bof then
	linkinfo = linkinfo & "<tr><td colspan=6><hr align=center size=1 color="&forum_body(27)&"></td></tr>"
	do while not rs.eof
	if i>6 then 
	linkinfo = linkinfo & "<tr><td width=""16%"">"
	i=1
	else
	linkinfo = linkinfo & "<td width=""16%"">"
	end if
	linkinfo = linkinfo & "<a href="&rs(2)&" target=_blank><img src="""&rs(3)&""" border=0 alt="""&rs(0)&"："&rs(1)&""" height=31 width=88></a>"
	rs.movenext
	linkinfo = linkinfo & "</td>"
	if i=6 then linkinfo = linkinfo & "</tr>"
	i=i+1
	loop
	end if
	set rs=nothing
	if isnull(linkinfo) or linkinfo="" then linkinfo=" "
	myCache.add linkinfo,dateadd("n",9999,now)
	response.write linkinfo
	end if

response.write "</table></TD></TR></table>"

'=========生日显示开始================
if cint(forum_setting(29))=1 then
	myCache.name="userbirthday"
	if myCache.valid then
	response.write myCache.value
	'mycache.makeEmpty
	else
	dim age
	dim birthuser
	dim foundbirth
	dim showbirthday
	dim birthNum,birthday
	foundbirth=false
	birthNum=0
	'on error resume next
	sql="select birthuser from config where active=1"
	set rs=conn.execute(sql)
	if not isnull(rs(0)) or rs(0)<>"" then
	birthuser=split(rs(0),"$")
	if ubound(birthuser)<3 then
		foundbirth=false
        elseif cint(birthuser(1))<>0 then
                foundbirth=true
	elseif datediff("d",birthuser(2),Now())>0 then
		foundbirth=false
	else
		foundbirth=true
	end if
	else
	foundbirth=false
	end if
	if not foundbirth then
		set rs=conn.execute("select username,birthday from [user] where birthday<>''  order by userid") 
		do while not rs.eof
		if   isdate(rs(1)) then
                   if month(rs(1))=month(Now()) and  day(rs(1))=day(Now()) then
			age=datediff("yyyy",rs(1),Now())
			birthday=birthday & "<a href=dispuser.asp?name="&rs(0)&" title=祝"&age&"岁生日快乐！ target=_blank>〖祝 "&rs(0)&" 生日快乐<img src=pic/birthday00.gif align=absmiddle border=0>〗  </a> ,"
			birthNum=birthNum+1
                   end if
		end if
		rs.movenext
		loop
		rs.close
		set rs=nothing
             if   birthday="" then  birthday="今天没有朋友过生日..."
		conn.execute("update config set birthuser='" & birthday & "$" & birthNum & "$" & Now() & "' where active=1")
	else
		birthday=birthuser(0)
		birthNum=birthuser(1)
	end if
	showbirthday="<BR><table cellpadding=3 cellspacing=1 align=center class=tableborder1><TR ><Th align=left >≡ 今天过生日的用户（共 "&birthNum&" 人）≡</Th></TR><TR><TD width=100% vAlign=middle class=tablebody1  >"&birthday&"</TD></TR></table>"
	myCache.add showbirthday,date+1
	end if
end if
'=========生日显示结束================

response.write "<BR><table cellpadding=5 cellspacing=1 class=tableborder1 align=center style=""word-break:break-all;"">"&_
	"<TR><Th align=left colSpan=2 height=25>-=> 用户来访信息</Th></TR>"&_
	"<TR><TD vAlign=top class=tablebody1 height=25 width=100% >"

	Dim userip,userip2
	userip = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
	userip2 = Request.ServerVariables("REMOTE_ADDR")
	if userip = ""  then
		response.write "您的真实ＩＰ 是："&userip2&"，"
	else
		response.write "您的真实ＩＰ 是："&userip&"，"
	end if
	response.write usersysinfo(Request.ServerVariables("HTTP_USER_AGENT"),2)&"，"&usersysinfo(Request.ServerVariables("HTTP_USER_AGENT"),1)

response.write "</TD></TR>"&_
	"<TR><Th colSpan=2 align=left id=tabletitlelink height=25 style=""font-weight:normal""><b>-=> 论坛在线统计</b>&nbsp;"

	if request("action")="show" then
		response.write "[<a href=index.asp?action=off>关闭详细列表</a>]"
	else
		if cint(Forum_Setting(14))=1 and request("action")<>"off" then
		response.write "[<a href=index.asp?action=off>关闭详细列表</a>]"
		else
		response.write "[<a href=index.asp?action=show>显示详细列表</a>]"
		end if
	end if

response.write "[<a href=boardstat.asp?reaction=online>查看在线用户位置</a>]</Th></TR>"&_
	"<TR>"&_
	"<TD width=100% vAlign=top class=tablebody1>  目前论坛上总共有 <b>"&Forum_OnlineNum&"</b> 人在线，其中注册会员 <b>"&OnlineNum&"</b> 人，访客 <b>"&GuestNum&"</b> 人。<br>"&_
	"历史最高在线纪录是 <b>"&Maxonline&"</b> 人同时在线，发生时间是："&formatdatetime(MaxonlineDate,1)&"&nbsp;"&formatdatetime(MaxonlineDate,4)&""&_
	"<BR><font color="&Forum_body(8)&">名单图例</font>："

response.write "<img src="&Forum_info(7)&Forum_pic(0)&"> 总坛主   ‖　<img src="&Forum_info(7)&Forum_pic(1)&"> 论坛坛主   ‖   <img src="&Forum_info(7)&Forum_pic(2)&"> 论坛贵宾   ‖   <img src="&Forum_info(7)&Forum_pic(3)&"> 普通会员   ‖   <img src="&Forum_info(7)&Forum_pic(4)&"> 客人或隐身会员"

response.write "<hr size=1 color="&forum_body(27)&">"&_
	"<table width=100% border=0 cellspacing=0 cellpadding=0>"

	if request("action")="off" then
		call onlineuser(0,0,0)
	elseif request("action")="show" then
		call onlineuser(1,1,0)
	else
		call onlineuser(Forum_Setting(14),Forum_Setting(15),0)
	end if
%>
</table></TD></TR>
</TABLE><BR>
<table cellspacing=1 cellpadding=3 width="<%=Forum_body(12)%>" border=0 align=center>
<tr><td align=center><img src="<%=Forum_info(7)&Forum_pic(6)%>" align="absmiddle">&nbsp;没有新的帖子&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="<%=Forum_info(7)&Forum_pic(7)%>" align="absmiddle">&nbsp;有新的帖子&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="<%=Forum_info(7)&Forum_pic(14)%>" align="absmiddle">&nbsp;被锁定的论坛</td></tr>
</table>
<%
if instr(scriptname,"index.asp")>0 or instr(scriptname,"list.asp")>0 then
	if Forum_ads(2)=1 then
	call admove()
	end if
	if Forum_ads(13)=1 then
	call fixup()
	end if
end if

sub FastLogin()
%>
<table cellspacing=1 cellpadding=3 align=center class=tableBorder1>
<form action=login.asp?action=chk method=post>
<tr>
<th align=left id=tabletitlelink height=25 style="font-weight:normal">
<b>-=> 快速登录入口</b>
[<a href=reg.asp>注册用户</a>]　[<a href=lostpass.asp style="CURSOR: help">忘记密码</a>]
</th>
</tr>
<tr>
<td class=tablebody1 height=40 width="100%">
&nbsp;用户名：<input maxLength=16 name=username size=12>　　密码：<input maxLength=20 name=password size=12 type=password>　　<select name=CookieDate><option selected value=0>不保存</option><option value=1>保存一天</option><option value=2>保存一月</option><option value=3>保存一年</option></select>　　<input type=submit name=submit value="登 陆">
</td>
</tr>
</form>
</table>
<%
end sub

call footer()
%>
<!--#include file="online_l.asp"-->
<!--#include file="inc/ad_fixup.asp"-->