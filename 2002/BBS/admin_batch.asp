<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
stats="帖子批量管理"
dim rootid
dim id
dim Lasttopic,Lastpost
dim lastrootid,lastpostuser
dim ip,url
dim title,content
dim totalusetable
dim ars
ip=Request.ServerVariables("REMOTE_ADDR")
if not founduser then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>请登陆后进行操作。"
end if

if BoardID="" or (not isInteger(BoardID)) or BoardID="0" then
	Errmsg=Errmsg+"<br>"+"<li>错误的版面参数！请确认您是从有效的连接进入。"
	founderr=true
else
	BoardID=clng(BoardID)
end if

Dim UpdateBoardID,UpdateBoardID_1
UpdateBoardID=BoardParentStr & "," & BoardID
if request.form("announceid")="" then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>请指定相关帖子。"
end if
if not founderr then
	dim canbatchtopic
	canbatchtopic=false
	if (master or superboardmaster or boardmaster) and Cint(GroupSetting(45))=1 then
		canbatchtopic=true
	else
		canbatchtopic=false
	end if
	if FoundUserPer and Cint(GroupSetting(45))=1 then
		canbatchtopic=true
	elseif FoundUserPer and Cint(GroupSetting(45))=0 then
		canbatchtopic=false
	end if
	if not canbatchtopic then
		Errmsg=Errmsg+"<br><li>您没有执行此操作的权限。"
		founderr=true
	end if
end if
call nav()
if founderr then
	call head_var(2,0,"","")
	call dvbbs_error()
else
	call head_var(1,BoardDepth,0,0)
	select case request("action")
	case "lock"
		call lock()
	case "dele"
		call delete()
	case "move"
		call Tmove()
	case "istop"
		call istop()
	case "isbest"
		call isbest()
	case else
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>请选择相关参数。"
	end select
	if founderr then call dvbbs_error()
end if
call footer()


sub lock()
Dim id
for i=1 to request.form("Announceid").count
	ID=replace(request.form("Announceid")(i),"'","")
	conn.execute("update topic set LockTopic=1 where BoardID="&BoardID&" And topicid=" & ID)
next
sql="insert into log (l_announceid,l_boardid,l_touser,l_username,l_content,l_ip) values ("&id&","&boardid&",'more','"&membername&"','批量锁定','"&ip&"')"
conn.execute(sql)
sucmsg="<li>帖子批量锁定成功。<li>您的操作信息已经记录在案。"
call dvbbs_suc()
end sub

sub delete()
Dim id
for i=1 to request.form("Announceid").count
	ID=replace(request.form("Announceid")(i),"'","")
	set rs=conn.execute("select PostTable from topic where boardid="&boardid&" and topicid="&id)
	totalusetable=rs(0)
	set rs=conn.execute("select PostUserID,parentid from "&totalusetable&" where BoardID="&BoardID&" And rootid=" & ID)
	do while not rs.eof
	conn.execute("update [user] set userWealth=userWealth-"&Forum_user(3)&",userCP=userCP-"&Forum_user(13)&",userEP=userEP-"&Forum_user(8)&" where userid="&rs(0))
	if rs(1)=0 then
	call BoardNumSub(boardid,1,1)
	call AllboardNumSub(1,1)
	else
	call BoardNumSub(boardid,0,1)
	call AllboardNumSub(1,0)
	end if
	rs.movenext
	loop
	conn.execute("update "&totalusetable&" set LockTopic=2 where BoardID="&BoardID&" And rootid=" & ID)
	conn.execute("update topic set locktopic=2 where boardid="&boardid&" and topicid="&id)
next
call LastCount(boardid)
sql="insert into log (l_announceid,l_boardid,l_touser,l_username,l_content,l_ip) values ("&id&","&boardid&",'more','"&membername&"','批量删除','"&ip&"')"
conn.execute(sql)
boardtoday(boardid)
alltodays()
sucmsg="<li>帖子批量删除成功。<li>您的操作信息已经记录在案。"
call dvbbs_suc()
set rs=nothing
end sub

sub Tmove()
Dim id,newboard,trs
if request.form("newboard")="" or isnull(request.form("newboard")) or not isnumeric(request.form("newboard")) then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>如果您是批量转移帖子请选择相关论坛。"
	exit sub
elseif Cint(request.form("newboard"))=Cint(boardid) then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>请不要选择相同的论坛进行转移操作。"
	exit sub
else
	newboard=request.form("newboard")
end if
for i=1 to request.form("Announceid").count
	ID=replace(request.form("Announceid")(i),"'","")
	set rs=conn.execute("select PostTable,isbest from topic where boardid="&boardid&" and topicid="&id)
	if not(rs.eof and rs.bof) then
	totalusetable=rs(0)
	set ars=conn.execute("select count(*) from "&totalusetable&" where BoardID="&BoardID&" And rootid=" & ID)
	call BoardNumSub(boardid,1,ars(0))
	set trs=conn.execute("select ParentStr from board where boardid="&newboard)
	UpdateBoardID_1=trs(0) & "," & newboard
	UpdateBoardID=trs(0) & "," & newboard
	call BoardNumAdd(newboard,1,ars(0))
	conn.execute("update "&totalusetable&" set boardid="&newboard&" where BoardID="&BoardID&" And rootid=" & ID)
	conn.execute("update topic set boardid="&newboard&" where boardid="&boardid&" and topicid="&id)
	if rs(1)=1 then
	conn.execute("update besttopic set boardid="&newboard&" where BoardID="&BoardID&" And rootid=" & ID)
	end if
	end if
next
UpdateBoardID=BoardParentStr & "," & BoardID
call LastCount(boardid)
boardtoday(boardid)
'UpdateBoardID=UpdateBoardID_1
call LastCount(newboard)
boardtoday(newboard)
alltodays()
sql="insert into log (l_announceid,l_boardid,l_touser,l_username,l_content,l_ip) values ("&id&","&boardid&",'more','"&membername&"','批量移动','"&ip&"')"
conn.execute(sql)
sucmsg="<li>帖子批量移动成功。<li>您的操作信息已经记录在案。"
call dvbbs_suc()
set rs=nothing
end sub

sub istop()
Dim id
for i=1 to request.form("Announceid").count
	ID=replace(request.form("Announceid")(i),"'","")
	conn.execute("update topic set istop=1 where boardid="&boardid&" and topicid="&id)
next
sql="insert into log (l_announceid,l_boardid,l_touser,l_username,l_content,l_ip) values ("&id&","&boardid&",'more','"&membername&"','批量固顶','"&ip&"')"
conn.execute(sql)

sucmsg="<li>帖子批量固定成功。<li>您的操作信息已经记录在案。"
call dvbbs_suc()
end sub

sub isbest()
Dim id
for i=1 to request.form("Announceid").count
	ID=replace(request.form("Announceid")(i),"'","")
	set rs=conn.execute("select PostTable from topic where boardid="&boardid&" and topicid="&id)
	totalusetable=rs(0)
	set rs=conn.execute("select top 1 * from "&TotalUseTable&" where rootid="&id&" order by Announceid")
	if not (rs.eof and rs.bof) then
	sql="insert into bestTopic (title,boardid,Announceid,rootid,postusername,postuserid,dateandtime,expression) values ('"&checkstr(rs("topic"))&"',"&rs("boardid")&","&rs("Announceid")&","&rs("rootid")&",'"&checkstr(rs("username"))&"',"&rs("postuserid")&",'"&rs("dateandtime")&"','"&rs("expression")&"')"
	conn.execute(sql)
	conn.execute("update "&TotalUseTable&" set isbest=1 where Announceid=" & rs("Announceid"))
	conn.execute("update topic set isbest=1 where topicid="&id)
	end if
next
sql="insert into log (l_announceid,l_boardid,l_touser,l_username,l_content,l_ip) values ("&id&","&boardid&",'more','"&membername&"','批量精华','"&ip&"')"
conn.execute(sql)
sucmsg="<li>帖子批量精华成功。<li>您的操作信息已经记录在案。"
call dvbbs_suc()
set rs=nothing
end sub

'更新指定论坛信息
function LastCount(boardid)
Dim LastTopic,body,LastRootid,LastPostTime,LastPostUser
Dim LastPost,uploadpic_n,Lastpostuserid,Lastid
set rs=conn.execute("select top 1 T.title,b.Announceid,b.dateandtime,b.username,b.postuserid,b.rootid from "&NowUseBBS&" b inner join Topic T on b.rootid=T.TopicID where b.boardid="&boardid&" and  b.locktopic<2 order by b.announceid desc")
if not(rs.eof and rs.bof) then
	Lasttopic=replace(left(rs(0),15),"$","")
	LastRootid=rs(1)
	LastPostTime=rs(2)
	LastPostUser=rs(3)
	LastPostUserid=rs(4)
	Lastid=rs(5)
else
	LastTopic="无"
	LastRootid=0
	LastPostTime=now()
	LastPostUser="无"
	LastPostUserid=0
	Lastid=0
end if
set rs=nothing

LastPost=LastPostUser & "$" & LastRootid & "$" & LastPostTime & "$" & LastTopic & "$" & uploadpic_n & "$" & LastPostUserID & "$" & LastID & "$" & BoardID
Dim SplitUpBoardID,SplitLastPost
SplitUpBoardID=split(UpdateBoardID,",")
For i=0 to ubound(SplitUpBoardID)
	set rs=conn.execute("select LastPost from board where boardid="&SplitUpBoardID(i))
	if not (rs.eof and rs.bof) then
	SplitLastPost=split(rs(0),"$")
	if SplitLastPost(1)="" or isnull(SplitLastPost(1)) then SplitLastPost(1)=0
	if ubound(SplitLastPost)=7 and clng(LastRootID)<>clng(SplitLastPost(1)) then
		conn.execute("update board set LastPost='"&LastPost&"' where boardid="&SplitUpBoardID(i))
	end if
	end if
Next
set rs=nothing
'sql="update board set LastPost='"&LastPost&"' where boardid in ("&UpdateBoardID&")"
'conn.execute(sql)
end function
	
'版面发帖数增加
sub BoardNumAdd(boardid,topicNum,postNum)
sql="update board set lastbbsnum=lastbbsnum+"&postNum&",lasttopicNum=lasttopicNum+"&topicNum&" where boardid in ("&UpdateBoardID&")"
conn.execute(sql)
end sub
	
'版面发帖数减少
sub BoardNumSub(boardid,topicNum,postNum)
sql="update board set lastbbsnum=lastbbsnum-"&postNum&",lasttopicNum=lasttopicNum-"&topicNum&" where boardid in ("&UpdateBoardID&")"
conn.execute(sql)
end sub
	
	
'所有论坛发帖数增加
function AllboardNumAdd(postNum,topicNum)
sql="update config set BbsNum=bbsNum+"&postNum&",TopicNum=topicNum+"&TopicNum
conn.execute(sql)
end function

'所有论坛发帖数减少
function AllboardNumSub(postNum,topicNum)
sql="update config set BbsNum=bbsNum-"&postNum&",TopicNum=topicNum-"&TopicNum
conn.execute(sql)
end function
'今日帖子
function boardtoday(boardid)
dim tmprs
tmprs=conn.execute("Select count(announceid) from "&NowUseBBS&" Where boardid in ("&UpdateBoardID&") and locktopic<2 and datediff('d',dateandtime,Now())=0")
boardtoday=tmprs(0)
set tmprs=nothing 
if isnull(boardtoday) then boardtoday=0
conn.execute("update board set todaynum="&boardtoday&" where boardid in ("&UpdateBoardID&")")
end function 

function alltodays()
dim tmprs
tmprs=conn.execute("Select count(announceid) from "&NowUseBBS&" Where locktopic<2 and datediff('d',dateandtime,Now())=0")
alltodays=tmprs(0)
set tmprs=nothing
if isnull(alltodays) then alltodays=0
conn.execute("update config set todaynum="&alltodays)
end function

%>