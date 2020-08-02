<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/char_board.asp" -->
<%
response.buffer=true
'=========================================================
' File: list.asp
' Version:5.0
' Date: 2002-9-20
' Script Written by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================

Dim currentPage,tl,limitime
Dim Fast_action
Dim reBoard_Setting,BoardCount,k,ColSpanNum
Dim Board_Count
Dim TdTableWidth
Dim master_3,master_4
dim onlinenum,guestnum
Dim LastPost,Lastuser,LastID
dim LastTime,LastUserid,LastRootid,body
dim Pnum,p,replynum
dim totalrec,ii,page_count
dim n,pi
dim call_info
dim RsTopic
dim cmd,AllTopNum,RsTop
AllTopNum=0
Fast_action=false

if boardmaster or master or superboardmaster then
	Fast_action=true
else
	Fast_action=false
end if

if BoardID="" or (not isInteger(BoardID)) or BoardID="0" then
	Errmsg=Errmsg+"<br>"+"<li>错误的版面参数！请确认您是从有效的连接进入。"
	founderr=true
else
	BoardID=clng(BoardID)
end if
if cint(Board_Setting(2))=1 then
	if not founduser then
		Errmsg=Errmsg+"<br>"+"<li>本论坛为认证论坛，请<a href=login.asp>登陆</a>并确认您的用户名已经得到管理员的认证后进入。"
		founderr=true
	else
		if chkboardlogin(boardid,membername)=false then
		Errmsg=Errmsg+"<br>"+"<li>本论坛为认证论坛，请确认您的用户名已经得到管理员的认证后进入。"
		founderr=true
		end if
	end if
end if
if cint(Board_Setting(1))=1 then
	if Cint(GroupSetting(37))=0 then
		Errmsg=ErrMsg+"<Br>"+"<li>您没有权限进入隐含论坛！"
		founderr=true
	end if
end if
if founderr then
	call nav()
	call head_var(2,0,"","")
	call dvbbs_error()
else
	call nav()
	if Cint(Board_Setting(43))=0 then
	stats="帖子列表"
	else
	stats="论坛列表"
	end if
	call head_var(1,BoardDepth,0,0)
	call main()
	if founderr then call dvbbs_error()
end if
call footer()

sub main()
currentPage=request("page")
if currentpage="" or not isInteger(currentpage) then
	currentpage=1
else
	currentpage=clng(currentpage)
end if
if request("selTimeLimit")="all" then
	tl=""
elseif request("selTimeLimit")="" then
	tl=""
else
	limitime=request("selTimeLimit")
	tl=" and dateandtime>='"&cstr(cdate(now()-limitime))&"' "
end if
stats=boardtype & "帖子列表"
call activeonline()
onlinenum=online(boardid)
guestnum=guest(boardid)
if trim(boardmasterlist)<>"" then
	master_1=split(boardmasterlist, "|")
	for i = 0 to ubound(master_1)
	master_2=""+master_2+"<a href=""dispuser.asp?name="+master_1(i)+""" target=_blank title=点击查看该版主资料>"+master_1(i)+"</a>&nbsp;"
	next
else
	master_2="无版主"
end if

response.write "<TABLE cellSpacing=0 cellPadding=0 width="&Forum_body(12)&" border=0 align=center>"&_
	"<tr>"&_
	"<td align=center width=100% valign=middle>"&_
	"<SCRIPT LANGUAGE='JavaScript' SRC='js/fader.js' TYPE='text/javascript'></script>"&_
	"<SCRIPT LANGUAGE='JavaScript' TYPE='text/javascript'>"&_
	"prefix="""";"&_
	"arNews = ["""
myCache.name="AnnounceMents"&BoardID
if myCache.valid then
	response.write myCache.value
else
Dim AnnounceMent_list
sql="select top 1 title,addtime from bbsnews where boardid="&boardid&"  order by id desc"
set rs=conn.execute(sql)
if rs.bof and rs.eof then
	AnnounceMent_list="<b><a href='announcements.asp?boardid="& boardid &"' target=_blank><ACRONYM TITLE='当前没有公告'>当前没有公告</ACRONYM></a></b> ("& now() &")"
else
	AnnounceMent_list="<b><a href='announcements.asp?boardid="& boardid &"' target=_blank>"& rs("title") &"</a></b> ("& rs("addtime") &")"
end if
set rs=nothing
myCache.add AnnounceMent_list,dateadd("n",60,now)
response.write AnnounceMent_list
end if
response.write ""","""",""欢迎光临"&Forum_info(0)&" "&boardtype&""","""",""严禁恶意使用粗言秽语，违者经劝告无效，立即封ID。"",""""]"&_
	"</SCRIPT>"&_
	"<div id=""elFader"" style=""position:relative;visibility:hidden; height:16"" ></div>"&_
	"</td>"&_
	"</tr>"&_
	"</table>"
%>
<script>
function loadThreadFollow(t_id,b_id){
	var targetImg =eval("document.all.followImg" + t_id);
	var targetDiv =eval("document.all.follow" + t_id);
	
	if ("object"==typeof(targetImg)){
		if (targetDiv.style.display!='block'){
			targetDiv.style.display="block";
			targetImg.src="<%=Forum_info(7)&Forum_boardpic(16)%>";
			if (targetImg.loaded=="no"){
				document.frames["hiddenframe"].location.replace("loadtree1.asp?boardid="+b_id+"&rootid="+t_id);
			}
		}else{
			targetDiv.style.display="none";
			targetImg.src="<%=Forum_info(7)&Forum_boardpic(15)%>";
		}
	}
}
</script>
<iframe width=0 height=0 src="" id="hiddenframe"></iframe>
<%
k=0
Dim cBoardID
dim simdisp
simdisp=false
Response.Cookies("BoardList").expires= date+7
cBoardID= Request("cBoardid")
if isnumeric(cBoardID) then
	if request("CatLog")="N" then
	Response.Cookies("BoardList")(cBoardID & "BoardID")= "NotShow"
	else
	Response.Cookies("BoardList")(cBoardID & "BoardID")= "Show"
	end if
end if
if request.cookies("BoardList")(boardid & "BoardID")<>"NNotShow" then
if BoardParentID=0 then
sql="select * from board where (depth<="&Boarddepth&"+"&forum_setting(5)&" and ParentStr like '%"&BoardID&"%') or boardid="&boardid&" order by orders"
else
sql="select * from board where (depth<="&Boarddepth&"+"&forum_setting(5)&" and ParentStr like '%"&BoardParentStr&","&boardid&"%') or boardid="&boardid&" order by orders"
end if
'response.write sql
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if not (rs.eof and rs.bof) then BoardCount=rs.recordcount
if BoardCount>1 then
do while not rs.eof
k=k+1
reBoard_Setting=split(rs("Board_setting"),",")
if rs("BoardMaster")<>"" and not isnull(rs("BoardMaster")) then
	master_3=split(rs("BoardMaster"), "|")
	for i = 0 to ubound(master_3)
	if i>6 then
		master_4=master_4
	else
		master_4=master_4 & "<a href=dispuser.asp?name="&master_3(i)&" target=_blank>"&master_3(i)&"</a>&nbsp;"
	end if
	next
	if i>7 then master_4=master_4 & "<font color=gray>More...</font>"
else
	master_4="暂无"
end if
if rs("ParentID")=0 or rs("boardid")=boardid then
	if request.cookies("BoardList")(rs("boardid") & "BoardID")="NotShow" then
		ColSpanNum=reBoard_Setting(41)
		simdisp=true
	elseif request.cookies("BoardList")(rs("boardid") & "BoardID")="Show" then
		simdisp=false
		ColSpanNum=2
	else
		if Cint(reBoard_Setting(39))=1 then
			ColSpanNum=reBoard_Setting(41)
			simdisp=true
		else
			simdisp=false
			ColSpanNum=2
		end if
	end if
	response.write "<table cellspacing=1 cellpadding=0 align=center class=tableBorder1>"
	response.write "<TR><Th colSpan="&ColSpanNum&" height=25 align=left id=TableTitleLink>&nbsp;"
	if simdisp then
		response.write "<a href=""?BoardID="&rs("boardid")&"&cBoardid="&rs("boardid")&"&Catlog=Y"" title=""展开论坛列表""><img src="""&Forum_info(7)&Forum_boardpic(15)&""" border=0></a><a href=list.asp?boardid="&rs("boardid")&" title=进入本分类论坛>"&rs("boardtype")&"</a>&nbsp;<a href=""?BoardID="&rs("boardid")&"&Catlog=NN"">[完全关闭]</a>"
	else
		response.write "<a href=""?BoardID="&rs("boardid")&"&cBoardid="&rs("boardid")&"&Catlog=N"" title=""关闭论坛列表""><img src="""&Forum_info(7)&Forum_boardpic(16)&""" border=0></a><a href=list.asp?boardid="&rs("boardid")&" title=进入本分类论坛>"&rs("boardtype")&"</a>&nbsp;<a href=""?BoardID="&rs("boardid")&"&Catlog=NN"">[完全关闭]</a>"
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
master_4=""
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
response.write "</table>"
end if
rs.close
set rs=nothing
response.write "<BR>"
end if
if currentpage=1 then
myCache.name="AllTopNum"
if myCache.valid then
	if myCache.value>0 then
	set RsTop=server.createobject("adodb.recordset")
	sql="select TopicID,boardid,title,postusername,postuserid,dateandtime,child,hits,votetotal,lastpost,lastposttime,istop,isvote,isbest,locktopic,Expression from topic where istop=2 and locktopic<2  ORDER BY lastposttime desc"
	RsTop.open sql,conn,1,1
	if RsTop.eof and RsTop.bof then
		AllTopNum=0
	else
		AllTopNum=RsTop.recordcount
	end if
	myCache.add AllTopNum,dateadd("n",60,now)
	end if 'end by value
else
	set RsTop=server.createobject("adodb.recordset")
	sql="select 	TopicID,boardid,title,postusername,postuserid,dateandtime,child,hits,votetotal,lastpost,lastposttime,istop,isvote,isbest,locktopic,Expression from topic where istop=2 and locktopic<2  ORDER BY lastposttime desc"
	RsTop.open sql,conn,1,1
	if RsTop.eof and RsTop.bof then
		AllTopNum=0
	else
		AllTopNum=RsTop.recordcount
	end if
	myCache.add AllTopNum,dateadd("n",60,now)
end if 'end by valid
end if 'end by page

if limitime="" then
	totalrec=lasttopicnum
else
	set rs=conn.execute("select count(topicid) from topic where boardID="&cstr(boardID)&" "&tl&"")
	totalrec=rs(0)
end if
set RsTopic=server.createobject("adodb.recordset")
if IsAudit=1 then
if not founduser then userid=0
sql="select TopicID,boardid,title,postusername,postuserid,dateandtime,child,hits,votetotal,lastpost,lastposttime,istop,isvote,isbest,locktopic,Expression from topic where boardID="&boardid&" and istop<2 and (locktopic<2 or (locktopic=3 and PostUserID="&userid&")) ORDER BY istop desc,lastposttime desc"
else
sql="select TopicID,boardid,title,postusername,postuserid,dateandtime,child,hits,votetotal,lastpost,lastposttime,istop,isvote,isbest,locktopic,Expression from topic where boardid="&boardid&" and istop<2 and locktopic<2 order by istop desc,lastposttime desc"
end if

RsTopic.open sql,conn,1
if RsTopic.eof and RsTopic.bof and Cint(Board_Setting(43))=0 and AllTopNum=0 then
	call BoardEmpty()
elseif not (RsTopic.eof and RsTopic.bof) or (Cint(Board_Setting(43))=0 and AllTopNum>0) then
	call listinfo()
end if
end sub

sub main_1_boardlist_1()

if (Cint(reBoard_Setting(1))=1 and Cint(GroupSetting(37))=1) or Cint(reBoard_Setting(1))=0 then
response.write "<TR><TD align=middle width=100% class=tablebody1><table width=""100%"" cellspacing=0 cellpadding=0>"
LastPostInfo=split(rs("LastPost"),"$")
if not isdate(LastPostInfo(2)) then LastPostInfo(2)=Now()
'=========================BoardInfo============================
response.write "<TR><TD align=middle width=46 class=tablebody1>"
if reBoard_Setting(0)=1 then
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

response.write "</TD>"&_
	"<TD width=1 bgcolor="&Forum_body(27)&">"&_
	"<TD vAlign=top width=* class=tablebody1>"&_
	"<TABLE cellSpacing=0 cellPadding=2 width=100% border=0>"&_
	"<tr><td class=tablebody1 width=*><a href=list.asp?boardid="&rs(0)&">"&_
	"<font color="&Forum_body(14)&">"&rs("boardtype")&"</font></a>"
if rs("child")>0 then
response.write "(<font title=""有"&rs("child")&"个下属论坛"">"&rs("child")&"</font>)"
end if
response.write "</td>"&_
	"<td width=40 rowspan=2 align=center class=tablebody1>"
if rs("indeximg")<>"" then
	response.write "<table align=left><tr><td><a href=list.asp?boardid="&rs("boardid")&"><img src="&rs("indeximg")&" align=top border=0></a></td><td width=20></td></tr></table>"
end if
response.write "</td><td width=200 rowspan=2 class=tablebody1>"
if Cint(reBoard_Setting(2))=1 then
response.write "认证论坛，请认证用户进入浏览"
else
response.write "主题：<a href=""dispbbs.asp?Boardid="&LastPostInfo(7)&"&ID="&LastPostInfo(6)&"&replyID="&LastPostInfo(1)&"&skin=1"">"&htmlencode(left(LastPostInfo(3),10))&"..</a><BR>作者：<a href=""dispuser.asp?id="&htmlencode(LastPostInfo(5))&""" target=_blank>"&htmlencode(LastPostInfo(0))&"</a><BR>日期：" & FormatDateTime(LastPostInfo(2),1) & "" & FormatDateTime(LastPostInfo(2),4)&"&nbsp;<a href=""dispbbs.asp?Boardid="&LastPostInfo(7)&"&ID="&LastPostInfo(6)&"&replyID="&LastPostInfo(1)&"&skin=1""><IMG border=0 src="""&Forum_info(7)&Forum_statePic(5)&""" title=""转到："&htmlencode(LastPostInfo(3))&"""></a>"
end if
response.write "</TD></TR><TR><TD width=*><img src="&Forum_info(7)&Forum_pic(8)&" align=middle> "
response.write rs("readme")
response.write"</TD>"

response.write "</TR><TR><TD class=tablebody2 height=20 width=*>版主："&master_4&""

response.write "</TD><td width=40 align=center class=tablebody2>&nbsp;</td><TD vAlign=middle class=tablebody2 width=200><table width=100% border=0><tr>"

response.write "<td width=25% vAlign=middle><img src="&Forum_info(7)&Forum_pic(9)&" alt=今日帖 align=absmiddle>&nbsp;<font color="&Forum_body(8)&">"&rs("TodayNum")&"</font></td><td width=30% vAlign=middle><img src="&Forum_info(7)&Forum_pic(11)&" alt=主题 border=0  align=absmiddle>&nbsp;"&rs("LastTopicNum")&"</td><td width=45% vAlign=middle><img src="&Forum_info(7)&Forum_pic(9)&" alt=文章 border=0 align=absmiddle>&nbsp;"&rs("LastBBSNum")&"</td></tr>"
response.write "</table></TD></TR></TBODY></TABLE></td></tr></table></td></tr>"
'============================End===============================
end if
end sub

sub main_1_boardlist_2()
TdTableWidth=100/ColSpanNum
if Board_Count=ColSpanNum or Board_Count=0 then response.write "<tr>"
response.write "<td class=tablebody1 width="""&TdTableWidth&"%"">"
	if (Cint(reBoard_Setting(1))=1 and Cint(GroupSetting(37))=1) or Cint(reBoard_Setting(1))=0 then
	response.write "<TABLE cellSpacing=2 cellPadding=2 width=100% border=0><tr><td width=""100%"" title="""&htmlencode(rs("readme"))&"<br>版主："&master_4&""" colspan=2><a href=list.asp?boardid="&rs("boardid")&"><font color="&Forum_body(14)&">"&rs("boardtype")&"</font></a>"
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

sub listinfo()
response.write "<TABLE cellpadding=3 cellspacing=1 class=tableborder1 align=center>"&_
	"<TR>"&_
	"<Th height=25 width=100% align=left id=tabletitlelink style=""font-weight:normal"">总在线<b>"&Forum_OnlineNum&"</b>人，其中"& boardtype &"上共有 <b>"& onlinenum &"</b> 位会员与 <b>"& guestnum &"</b> 位客人.今日贴子 <b><font color="&Forum_body(8)&">"& todaynum &"</font></b> "
if request("action")="show" then
	response.write "[<a href=list.asp?action=off&boardID="& boardid&"&amp;page=1&skin="& skin &">关闭详细列表</a>]"
else
	if cint(Forum_Setting(14))=1 and request("action")<>"off" then
	response.write "[<a href=list.asp?action=off&boardID="& boardid&"&amp;page=1&skin="& skin &">关闭详细列表</a>]"
	else
	response.write "[<a href=list.asp?action=show&boardID="& boardid&"&amp;page=1&skin="& skin &">显示详细列表</a>]"
	end if
end if

response.write "</Th></TR>"

if request("action")="off" then
	call onlineuser(0,0,boardid)
elseif request("action")="show" then
	call onlineuser(1,1,boardid)
else
	call onlineuser(Forum_Setting(14),Forum_Setting(15),boardid)
end if

response.write "</td></tr></TABLE><BR>"

response.write "<table cellpadding=0 cellspacing=0 border=0 width="&Forum_body(12)&" align=center valign=middle><tr>"&_
	"<td align=center width=2> </td>"&_
	"<td align=left>"
if (Cint(Board_Setting(43))=0 and Cint(Board_Setting(0))=0) or (Cint(Board_Setting(43))=0 and Cint(Board_Setting(0))=1 and (master or superboardmaster or boardmaster)) then
response.write "<a href=announce.asp?boardid="& boardid &"><img src="&Forum_info(7)&Forum_boardpic(1)&" border=0 alt=发新帖></a>"&_
	"&nbsp;&nbsp;<a href=vote.asp?boardid="&boardid&"><img src="&Forum_info(7)&Forum_boardpic(2)&" border=0 alt=发起新投票></a>"&_
	"&nbsp;&nbsp;<a href=smallpaper.asp?boardid="&boardid&"><img src="&Forum_info(7)&Forum_boardpic(3)&" border=0 alt=发布小字报></a>"
elseif Cint(Board_Setting(0))=1 then
response.write "<B>本论坛已锁定，不允许进行发贴/回帖等操作</B>"
end if
response.write "</td>"&_
	"<td align=right><img src="&Forum_info(7)&Forum_pic(13)&" align=absmiddle>  "& master_2 &"</td></tr></table>"

response.write "<TABLE cellPadding=1 cellSpacing=1 class=tableborder1 align=center>"&_
	"<tr><td class=tablebody1 colspan=5 height=20>"&_
	"<table width=100% >"&_
	"<tr>"&_
	"<td valign=middle height=20 width=50> <a href=AllPaper.asp?boardid="&boardid&" title=点击查看本论坛所有小字报><b>广播：</b></a> </td><td width=*> "&_
	"<marquee scrolldelay=150 scrollamount=4 onmouseout=""if (document.all!=null){this.start()}"" onmouseover=""if (document.all!=null){this.stop()}"">"
set rs=conn.execute("SELECT TOP 5 s_id,s_username,s_title FROM smallpaper where datediff('d',s_addtime,Now())<=1 and s_boardid="&boardid&" ORDER BY s_addtime desc")
do while not rs.eof
	response.write "<font color="&Forum_pic(5)&">"&htmlencode(rs(1))&"</font>说：<a href=javascript:openScript('viewpaper.asp?id="&rs(0)&"&boardid="&boardid&"',500,400)>"&htmlencode(rs(2))&"</a>"
rs.movenext
loop
set rs=nothing

response.write "</marquee><td align=right width=240><a href=elist.asp?boardid="& boardid &" title=查看本版精华><font color="&Forum_body(8)&"><B>精华</B></font></a> | <a href=boardstat.asp?reaction=online&boardid="&boardid&" title=查看本版在线详细情况>在线</a> | <a href=bbseven.asp?boardid="&boardid&" title=查看本版事件>事件</a> | <a href=BoardPermission.asp?boardid="&boardid&" title=查看本版用户组权限>权限</a> | <a href=admin_boardset.asp?boardid="&boardid&">管理</a>"
if isaudit=1 then
dim allaudit
set rs=conn.execute("select count(*) from "&Nowusebbs&" where boardid="&boardid&" and locktopic=3")
allaudit=rs(0)
if isnull(allaudit) then allaudit=0
set rs=nothing
response.write " | <a href=admin_topiclist.asp?boardid="&boardid&" title=有"&allaudit&"个帖子待审核>审核</a>(<font color="&Forum_body(8)&">"&allaudit&"</font>)"
end if
response.write "</td></tr></table>"&_
	"</td></tr>"&_
	"<form action=admin_batch.asp method=post name=batch>"&_
	"<TR align=middle>"&_
	"<Th height=25 width=32 id=tabletitlelink><a href=list.asp?boardid="&boardid&"&page="&request("page")&"&action=batch>状态</a></th>"&_
	"<Th width=* id=tabletitlelink>主 题  (点<img src="&Forum_info(7)&Forum_boardpic(15)&" align=absmiddle>即可展开贴子列表)</Th>"&_
	"<Th width=80 id=tabletitlelink>作 者</Th>"&_
	"<Th width=64 id=tabletitlelink>回复/人气</Th>"&_
	"<Th width=195 id=tabletitlelink>最后更新 | 回复人</Th>"&_
	"</TR>"
if RsTopic.bof and RsTopic.eof and AllTopNum=0 then
	response.write "<tr><td colSpan=5 width=100% class=tablebody1 height=25>本论坛暂无内容，欢迎发贴：）</td></tr>"
else
'=========================All Top=========================
if currentpage=1 and alltopnum>0 then
	do while not RsTop.eof
	page_count=page_count+1
response.write "<TR align=middle><TD class=tablebody2 width=32 height=27>"
'response.write RsTop("lastposttime")
response.write "<img src="""&Forum_info(7)&Forum_statePic(9)&""" alt=总固顶主题>"

response.write "</TD><TD align=left class=tablebody1 width=*>"

if RsTop(6)=0 then
	response.write "<img src="""& Forum_info(7)&Forum_boardpic(15)&""" id=""followImg"& RsTop("topicid") &""">"
else
	response.write "<img loaded=""no"" src="""& Forum_info(7)&Forum_boardpic(16)&""" id=""followImg"& RsTop("topicid") &""" style=""cursor:hand;"" onclick=""loadThreadFollow("& RsTop("topicid") &","& RsTop("boardid") &")"" title=展开贴子列表>"
end if

if not isnull(RsTop("lastpost")) then
	LastPost=split(RsTop("lastpost"),"$")
	if ubound(LastPost)=7 then
	Lastuser=htmlencode(LastPost(0))
	LastID=LastPost(1)
	LastTime=LastPost(2)
	if not isdate(LastTime) then LastTime=RsTop("dateandtime")
	body=htmlencode(LastPost(3))
	if not isnull(LastPost(4)) and LastPost(4)<>"" then response.write "<img src=""images/files/"&LastPost(4)&".gif"" >"
	LastUserid=LastPost(5)
	LastRootid=LastPost(6)
	else
	Lastuser="------"
	LastID=RsTop("topicid")
	LastTime=RsTop("dateandtime")
	body="..."
	LastUserid=RsTop("postuserid")
	LastRootid=RsTop("topicid")
	end if
else
	Lastuser="------"
	LastID=RsTop("topicid")
	LastTime=RsTop("dateandtime")
	body="..."
	LastUserid=RsTop("postuserid")
	LastRootid=RsTop("topicid")
end if
if LastTime=RsTop("dateandtime") then LastUser=""

response.write "<a href=""dispbbs.asp?boardID="& RsTop("boardid") &"&ID="& RsTop("topicid") &""" title=""《"& htmlencode(RsTop("title")) &"》<br>作者："& htmlencode(RsTop("postusername")) &"<br>发表于"& RsTop("dateandtime") &"<br>最后跟贴："& body &"..."">"

if len(RsTop("title"))>26 then
	response.write left(htmlencode(RsTop("title")),26)&"..."
else
	response.write htmlencode(RsTop("title"))
end if
response.write "</a>"
replynum=RsTop("child")+1
if replynum>Cint(Board_Setting(27)) then
	response.write "&nbsp;&nbsp;[<img src="&Forum_info(7)&Forum_statePic(6)&"><b>"
  	if replynum mod Cint(Board_Setting(27))=0 then
     		Pnum= replynum \ Cint(Board_Setting(27))
  	else
     		Pnum= replynum \ Cint(Board_Setting(27))+1
  	end if
	for p=1 to Pnum
	response.write " <a href='dispbbs.asp?boardID="& RsTop("boardID") &"&ID="& RsTop("topicid") &"&star="&P&"'><font color="&Forum_body(8)&">"&p&"</font></a> "
	if p+1>7 then
	response.write "... <a href='dispbbs.asp?boardID="& boardID &"&ID="& RsTop("topicid") &"&star="&Pnum&"'><font color="&Forum_body(8)&">"&Pnum&"</font></a>"
	exit for
	end if
	next
	response.write "</b>]"
else
	pnum=1
end if

response.write "</TD>"
response.write "<TD class=tablebody2 width=80><a href=""dispuser.asp?id="& RsTop("postuserid") &""" target=_blank>"& htmlencode(RsTop("postusername")) &"</a></TD>"
response.write "<TD class=tablebody1 width=64>"

if RsTop("isvote")=1 then
	response.write "<FONT color="&Forum_body(8)&"><b>"&RsTop("votetotal")&"</b></font>  票"
else
	response.write RsTop("child") &"/"& RsTop("hits")
end if

response.write "</TD>"
response.write "<TD align=left class=tablebody2 width=195>&nbsp;&nbsp;<a href=""dispbbs.asp?boardid="& RsTop("boardid") &"&id="&LastRootID&"&star="&pnum&"#"& LastID&""">"
response.write FormatDateTime(LastTime,2)&"&nbsp;"&FormatDateTime(LastTime,4)&"</a>&nbsp;<font color="&Forum_body(8)&">|</font>&nbsp;"

if LastUser="" then
	response.write "------"
else
	response.write "<a href=dispuser.asp?id="&LastUserID&" target=_blank>"&LastUser&"</a>"
end if

response.write "</TD></TR>"
response.write "<tr style=""display:none"" id=""follow"& RsTop("topicid") &"""><td colspan=5 id=""followTd"& RsTop("topicid") &""" style=""padding:0px""><div style=""width:240px;margin-left:18px;border:1px solid black;background-color:lightyellow;color:black;padding:2px"" onclick=""loadThreadFollow("& RsTop("topicid") &","&RsTop("boardid")&")"">正在读取关于本主题的跟贴，请稍侯……</div></td></tr>"
	Lastuser=""
	LastID=""
	LastTime=""
	body=""
	RsTop.movenext
	loop
	set RsTop=nothing
end if
'============================end==========================
	if totalrec mod Cint(Board_Setting(26))=0 then
		n= totalrec \ Cint(Board_Setting(26))
	else
	     	n= totalrec \ Cint(Board_Setting(26))+1
  	end if
  	if not (RsTopic.bof and RsTopic.eof) then
	RsTopic.MoveFirst
	if currentpage > n then currentpage = n
	if currentpage<1 then currentpage=1
	RsTopic.Move (currentpage-1) * Cint(Board_Setting(26))
	do while not RsTopic.eof and page_count<Cint(Board_Setting(26))
	page_count=page_count+1
'========================TopicInfo=========================
response.write "<TR align=middle><TD class=tablebody2 width=32 height=27>"
'response.write rstopic("lastposttime")
if request("action")="batch" and Cint(GroupSetting(45))=1 then
	response.write "<input type=checkbox name=Announceid value="&RsTopic("topicid")&">"
else
	if RsTopic("istop")=2 then
	response.write "<img src="""&Forum_info(7)&Forum_statePic(9)&""" alt=总固顶主题>"
	elseif RsTopic("istop")=1 then
	response.write "<img src="""&Forum_info(7)&Forum_statePic(3)&""" alt=固顶主题>"
	elseif RsTopic("isbest")=1 then
	response.write "<img src="""&Forum_info(7)&Forum_statePic(4)&""" alt=精华帖子>"
	elseif RsTopic("isvote")=1 then
	response.write "<img src="""&Forum_info(7)&Forum_statePic(12)&""" alt=投票贴子>"
	elseif RsTopic("locktopic")=1 then
	response.write "<img src="""&Forum_info(7)&Forum_statePic(2)&""" alt=本主题已锁定>"
	elseif RsTopic("child")>=Cint(forum_setting(44)) then
	response.write "<img src="""&Forum_info(7)&Forum_statePic(1)&""" alt=热门主题>"
	else
	response.write "<img src="""&Forum_info(7)&Forum_statePic(0)&""" alt=开放主题>"
	end if
end if

response.write "</TD><TD align=left class=tablebody1 width=* >"

if RsTopic(6)=0 then
	response.write "<img src="""& Forum_info(7)&Forum_boardpic(16)&""" id=""followImg"& RsTopic("topicid") &""">"
else
	response.write "<img loaded=""no"" src="""& Forum_info(7)&Forum_boardpic(15)&""" id=""followImg"& RsTopic("topicid") &""" style=""cursor:hand;"" onclick=""loadThreadFollow("& RsTopic("topicid") &","& RsTopic("boardid") &")"" title=展开贴子列表>"
end if

if not isnull(RsTopic("lastpost")) then
	LastPost=split(RsTopic("lastpost"),"$")
	if ubound(LastPost)=7 then
	Lastuser=htmlencode(LastPost(0))
	LastID=LastPost(1)
	LastTime=LastPost(2)
	if not isdate(LastTime) then LastTime=RsTopic("dateandtime")
	body=htmlencode(LastPost(3))
	if not isnull(LastPost(4)) and LastPost(4)<>"" then response.write "<img src=""images/files/"&LastPost(4)&".gif"" >"
	LastUserid=LastPost(5)
	LastRootid=LastPost(6)
	else
	Lastuser="------"
	LastID=RsTopic("topicid")
	LastTime=RsTopic("dateandtime")
	body="..."
	LastUserid=RsTopic("postuserid")
	LastRootid=RsTopic("topicid")
	end if
else
	Lastuser="------"
	LastID=RsTopic("topicid")
	LastTime=RsTopic("dateandtime")
	body="..."
	LastUserid=RsTopic("postuserid")
	LastRootid=RsTopic("topicid")
end if
if LastTime=RsTopic("dateandtime") then LastUser=""

response.write "<a href=""dispbbs.asp?boardID="& RsTopic("boardid") &"&ID="& RsTopic("topicid") &""" title=""《"& htmlencode(RsTopic("title")) &"》<br>作者："& htmlencode(RsTopic("postusername")) &"<br>发表于"& RsTopic("dateandtime") &"<br>最后跟贴："& body &"..."">"

if len(RsTopic("title"))>26 then
	response.write left(htmlencode(RsTopic("title")),26)&"..."
else
	response.write htmlencode(RsTopic("title"))
end if
response.write "</a>"
replynum=RsTopic("child")+1
if replynum>Cint(Board_Setting(27)) then
	response.write "&nbsp;&nbsp;[<img src="&Forum_info(7)&Forum_statePic(6)&"><b>"
  	if replynum mod Cint(Board_Setting(27))=0 then
     		Pnum= replynum \ Cint(Board_Setting(27))
  	else
     		Pnum= replynum \ Cint(Board_Setting(27))+1
  	end if
	for p=1 to Pnum
	response.write " <a href='dispbbs.asp?boardID="& boardID &"&ID="& RsTopic("topicid") &"&star="&P&"'><font color="&Forum_body(8)&">"&p&"</font></a> "
	if p+1>7 then
	response.write "... <a href='dispbbs.asp?boardID="& boardID &"&ID="& RsTopic("topicid") &"&star="&Pnum&"'><font color="&Forum_body(8)&">"&Pnum&"</font></a>"
	exit for
	end if
	next
	response.write "</b>]"
else
	pnum=1
end if

response.write "</TD>"
response.write "<TD class=tablebody2 width=80><a href=""dispuser.asp?id="& RsTopic("postuserid") &""" target=_blank>"& htmlencode(RsTopic("postusername")) &"</a></TD>"
response.write "<TD class=tablebody1 width=64>"

if RsTopic("isvote")=1 then
	response.write "<FONT color="&Forum_body(8)&"><b>"&RsTopic("votetotal")&"</b></font>  票"
else
	response.write RsTopic("child") &"/"& RsTopic("hits")
end if

response.write "</TD>"
response.write "<TD align=left class=tablebody2 width=195>&nbsp;&nbsp;<a href=""dispbbs.asp?boardid="& boardid &"&id="&LastRootID&"&star="&pnum&"#"& LastID&""">"
response.write FormatDateTime(LastTime,2)&"&nbsp;"&FormatDateTime(LastTime,4)&"</a>&nbsp;<font color="&Forum_body(8)&">|</font>&nbsp;"

if LastUser="" then
	response.write "------"
else
	response.write "<a href=dispuser.asp?id="&LastUserID&" target=_blank>"&LastUser&"</a>"
end if

response.write "</TD></TR>"
response.write "<tr style=""display:none"" id=""follow"& RsTopic("topicid") &"""><td colspan=5 id=""followTd"& RsTopic("topicid") &""" style=""padding:0px""><div style=""width:240px;margin-left:18px;border:1px solid black;background-color:lightyellow;color:black;padding:2px"" onclick=""loadThreadFollow("& RsTopic("topicid") &","&RsTopic("boardid")&")"">正在读取关于本主题的跟贴，请稍侯……</div></td></tr>"
'===========================End============================
	Lastuser=""
	LastID=""
	LastTime=""
	body=""
	RsTopic.movenext
	loop
	end if
	set RsTopic=nothing
	if request("action")="batch" then

	response.write "<tr><td height=30 width=100% class=tablebody1 colspan=5><select name=newboard size=1>"
	response.write "<option selected>移动帖子请选择</option>"

	set rs=conn.execute("select boardid,boardtype,depth from board order by rootid,orders")
	do while not rs.EOF
	response.write "<option value="""&rs(0)&""" "
	response.write ">"
	select case rs(2)
	case 0
	response.write "╋"
	case 1
	response.write "&nbsp;&nbsp;├"
	end select
	if rs(2)>1 then
	for i=2 to rs(2)
		response.write "&nbsp;&nbsp;│"
	next
	response.write "&nbsp;&nbsp;├"
	end if
	response.write rs(1)&"</option>"
	rs.MoveNext
	loop
	set rs=nothing

response.write "<input type=hidden value="&boardid&" name=boardid><input type=checkbox name=chkall value=on onclick=""CheckAll(this.form)"">全选/取消 <input type=radio name=action value=dele>批量删除 <input type=radio name=action value=move>批量移动 <input type=radio name=action value=isbest>批量精华 <input type=radio name=action value=lock>批量锁定 <input type=radio name=action value=istop>批量固顶 <input type=submit name=Submit value=执行  onclick=""{if(confirm('您确定执行的操作吗?')){this.document.batch.submit();return true;}return false;}""></td></tr>"
	end if
	response.write "</form>"
end if
set rs=nothing
if totalrec mod Board_Setting(26)=0 then
   	n= totalrec \ Board_Setting(26)
else
   	n= totalrec \ Board_Setting(26)+1
end if
if currentpage-1 mod 10=0 then
	p=(currentpage-1) \ 10
else
	p=(currentpage-1) \ 10
end if
dim pagelist,pagelistbit

response.write "</table>"

response.write "<table border=0 cellpadding=0 cellspacing=3 width="&Forum_body(12)&" align=center >"&_
	"</form>"&_
	"<form method=post action=list.asp?boardid="&boardid&"&action="&request("action")&">"&_
	"<tr>"&_
	"<td valign=middle>页次：<b>"& currentPage &"</b>/<b>"& n &"</b>页 每页<b>"& Board_Setting(26) &"</b> 主题数<b>"& totalrec &"</b></td>"&_
	"<td valign=middle><div align=right >分页："

if currentPage=1 then
	response.write "<font face=webdings color="&Forum_body(8)&">9</font>   "
else
	response.write "<a href='?boardid="&boardid&"&page=1&action="&request("action")&"' title=首页><font face=webdings>9</font></a>   "
end if
if p*10>0 then response.write "<a href='?boardid="&boardid&"&page="&Cstr(p*10)&"&action="&request("action")&"' title=上十页><font face=webdings>7</font></a>   "
response.write "<b>"
for ii=p*10+1 to P*10+10
   if ii=currentPage then
		response.write "<font color="&Forum_body(8)&">"+Cstr(ii)+"</font> "
	else
		response.write "<a href='?boardid="&boardid&"&page="&Cstr(ii)&"&action="&request("action")&"'>"+Cstr(ii)+"</a>   "
	end if
	if ii=n then exit for
	'p=p+1
next
response.write "</b>"
if ii<n then response.write "<a href='?boardid="&boardid&"&page="&Cstr(ii)&"&action="&request("action")&"' title=下十页><font face=webdings>8</font></a>   "
if currentPage=n then
	response.write "<font face=webdings color="&Forum_body(8)&">:</font>   "
else
	response.write "<a href='?boardid="&boardid&"&page="&Cstr(n)&"&action="&request("action")&"' title=尾页><font face=webdings>:</font></a>   "
end if

response.write "转到:<input type=text name=Page size=3 maxlength=10  value="& currentpage &"><input type=submit value=Go name=submit></div></td></tr>"&_
	"</form></table>"

response.write "<table border=0 cellpadding=0 cellspacing=3 width="&Forum_body(12)&" align=center>"&_
	"<tr>"&_
	"<FORM METHOD=POST ACTION=queryresult.asp?boardid="&boardid&"&sType=2&SearchDate=30&pSearch=1>"&_
	"<td width=50% valign=middle nowrap height=40>"&_
	"快速搜索：<input type=text name=keyword>&nbsp;<input type=submit name=submit value=搜索>"&_
	"</td>"&_
	"</FORM>"&_
	"<td valign=middle nowrap width=50% > <div align=right>"&_
	"<select onchange=""if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"">"&_
	"<option selected>跳转论坛至...</option>"
myCache.name="BoardJumpList"
if myCache.valid then
response.write myCache.value
else
Dim BoardJumpList
set rs=conn.execute("select boardid,boardtype,depth from board order by rootid,orders")
do while not rs.EOF
BoardJumpList = BoardJumpList & "<option value=""list.asp?boardid="&rs(0)&""" "
BoardJumpList = BoardJumpList & ">"
select case rs(2)
case 0
BoardJumpList = BoardJumpList & "╋"
case 1
BoardJumpList = BoardJumpList & "&nbsp;&nbsp;├"
end select
if rs(2)>1 then
for i=2 to rs(2)
	BoardJumpList = BoardJumpList & "&nbsp;&nbsp;│"
next
BoardJumpList = BoardJumpList & "&nbsp;&nbsp;├"
end if
BoardJumpList = BoardJumpList & rs(1)&"</option>"
rs.MoveNext
loop
set rs=nothing
myCache.add BoardJumpList,dateadd("n",60,now)
response.write BoardJumpList
end if
response.write "</select><div></td></tr></table><BR>"

response.write "<table cellspacing=1 cellpadding=3 width=100% class=tableborder1 align=center><tr>"&_
	"<th width=80% align=left>　-=> "& Forum_info(0) &"图例</th>"&_
	"<th noWrap width=20% align=right>所有时间均为 - "&Forum_info(9)&" &nbsp;</th>"&_
	"</tr>"&_
	"<tr><td colspan=2 class=tablebody1>"&_
	"<table cellspacing=4 cellpadding=0 width=92% border=0 align=center>"&_
	"<tr><td><img src="&Forum_info(7)&Forum_statePic(0)&"> 开放的主题</td>"&_
	"<td><img src="&Forum_info(7)&Forum_statePic(1)&"> 回复超过"&forum_setting(44)&"贴</td>"&_
	"<td><img src="&Forum_info(7)&Forum_statePic(2)&"> 锁定的主题</td>"&_
	"<td><img src="&Forum_info(7)&Forum_statePic(3)&"> 固顶的主题 </td>"&_
	"<td><img src="&Forum_info(7)&Forum_statePic(9)&"> 总固顶的主题 </td>"&_
	"<td> <img src="&Forum_info(7)&Forum_statePic(4)&"> 精华帖子 </td>"&_
	"<td> <img src="&Forum_info(7)&Forum_statePic(12)&"> 投票帖子 </td>"&_
	"</tr>"&_
	"</table>"&_
	"</td></tr></table>"
end sub

sub BoardEmpty()
response.write "<table cellpadding=0 cellspacing=0 border=0 width="&Forum_body(12)&" align=center valign=middle><tr>"&_
	"<td align=center width=2> </td>"&_
	"<td align=left> <a href=""announce.asp?boardid="& boardid &"""><img src="&Forum_info(7)&Forum_boardpic(1)&" border=0 alt=发新帖></a>"&_
	"&nbsp;&nbsp;<a href=vote.asp?boardid="&boardid&"><img src="&Forum_info(7)&Forum_boardpic(2)&" border=0 alt=发起新投票></a>"&_
	"&nbsp;&nbsp;<a href=smallpaper.asp?boardid="&boardid&"><img src="&Forum_info(7)&Forum_boardpic(3)&" border=0 alt=发布小字报></a></td>"&_
	"<td align=right><img src="&Forum_info(7)&Forum_pic(13)&" align=absmiddle>  "& master_2 &"</td></tr></table>"&_
	"<TABLE cellPadding=1 cellSpacing=1 class=tableborder1 align=center>"&_
	"<TR align=middle>"&_
	"<Th height=25 width=32 id=tabletitlelink><a href=list.asp?boardid="&boardid&"&page="&request("page")&"&action=batch>状态</a></th>"&_
	"<Th width=*>主 题  (点<img src="&Forum_info(7)&Forum_boardpic(15)&">即可展开贴子列表)</Th>"&_
	"<Th width=80>作 者</Th>"&_
	"<Th width=64>回复/人气</Th>"&_
	"<Th width=195>最后更新 | 回复人</Th>"&_
	"</TR>"&_
	"<tr><td colSpan=5 width=100% class=tablebody1 height=25>本论坛暂无内容，欢迎发贴：）</td></tr></table><BR>"
end sub

'if instr(scriptname,"index.asp")>0 or instr(scriptname,"list.asp")>0 then
'if Forum_ads(2)=1 then
'call admove()
'end if
'if Forum_ads(13)=1 then
'call fixup()
'end if
'end if
%>
<!--#include file="online_l.asp"-->
<!--#include file="inc/ad_fixup.asp"-->