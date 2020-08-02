<!--#include file=char.asp-->
<%
'论坛Cookies名称，一般不需要修改，当您的网站启用两个以上的论坛时必须将这些论坛系统的Cookies名称修改为不同
Dim Forum_sn
Dim Mycache
set myCache=new Cache
Forum_sn="aspsky"
if Request.Cookies("iscookies")="" then
	Response.Cookies("iscookies")="0"
	Response.Cookies("iscookies").Expires=date+3650
	response.write "<META http-equiv=Content-Type content=text/html; charset=gb2312><meta HTTP-EQUIV=REFRESH CONTENT=3>正在登陆论坛……<br><br>本系统要求使用COOKIES，假如您的浏览器禁用COOKIES，您将不能登录本系统……"
	response.end
end if

Dim UserAgent,Stats,ScriptName
UserAgent=Trim(lcase(Request.Servervariables("HTTP_USER_AGENT")))
ScriptName=lcase(request.ServerVariables("PATH_INFO"))
If Instr(UserAgent,"teleport")>0 or Instr(UserAgent,"webzip")>0 or Instr(UserAgent,"flashget")>0 or Instr(UserAgent,"offline")>0 Then
	response.redirect "error.htm"
	response.end
end if

Dim copyright,Version,Cookiepath,Badwords,Splitwords,StopReadme,Maxonline,MaxonlineDate
Dim LastTopicNum,LastPostTime,LastPostInfo,lastbbsnum
Dim guestlist
Dim boardtype,lockboard,master_2,todaynum,master_1,boardmasterlist
Dim skinid,skin,tid,tempid
Dim templateused
Dim Forum_info,Forum_setting,Forum_ubb,Forum_body,Forum_ads,Forum_user,Board_Setting
Dim Forum_pic,Forum_boardpic,Forum_TopicPic,Forum_statePic,Forum_upload
Dim MaxOldID,NowUseBbs,AllPostTable,AllPostTableName,BoardReadme
Dim sqlstr
Dim mastermsg
Dim Forum_getre,Forum_smsmsg
Dim backpage,navinfo
Dim Forum_header
Dim membername,memberword,memberclass,userhidden
Dim sql,rs,ors
Dim BoardID,FoundBoard,i
Dim Founderr,Errmsg,sucmsg
Dim Founduser,userid
Dim FoundStyle
Dim UserPermission,FoundUserPer
Dim BoardParentStr,BoardParentID,BoardDepth,BoardChild
Dim Forum_OnlineNum
Dim Forum_userface,Forum_PostFace,Forum_emot
Dim Forum_userfaceNum,Forum_PostFaceNum,Forum_emotNum
Dim OpenTime,isaudit
Forum_OnlineNum=allonline()
FoundUserPer=false
FoundBoard=false
Founduser=false
Founderr=false
FoundStyle=false
if request("BoardID")="" or (not isInteger(request("BoardID"))) or request("boardid")="0" or instr(scriptname,"index.asp")>0 then
	BoardID=0
	FoundBoard=false
else
	BoardID=Clng(Request("BoardID"))
	FoundBoard=true
end if
membername=checkStr(request.cookies("aspsky")("username"))
memberword=checkStr(request.cookies("aspsky")("password"))
memberclass=checkStr(request.cookies("aspsky")("userclass"))
userhidden=request.cookies("aspsky")("userhidden")
userid=request.cookies("aspsky")("userid")
if not isnumeric(userhidden) or userhidden="" then userhidden=2
'论坛基本变量调用
if not FoundBoard then
	skinid=request.cookies("skin")("skinid_0")
	if isnumeric(skinid) and not isnull(skinid) and skinid<>"" then
		sqlstr=" id="&skinid&" "
	else
		if not isnumeric(request("skinid")) or request("skinid")="" or request("skinid")="0" then
			sqlstr=" active=1"
		else
			sqlstr=" id="&request("skinid")&" "
		end if
	end if
	sql = "select * from config where "&sqlstr&""
else
	skinid=request.cookies("skin")("skinid_"&boardid)
	if isnumeric(skinid) and not isnull(skinid) and skinid<>"" then
	sql = "select b.boardtype,b.boardmaster,b.todaynum,b.lasttopicnum,b.lastbbsnum,b.LastPost,b.readme,b.Board_Setting,b.ParentID,b.Depth,b.ParentStr,b.child,c.* from board b,config c where c.id="&skinid&" and  b.boardid="&boardid
	else
	sql = "select b.boardtype,b.boardmaster,b.todaynum,b.lasttopicnum,b.lastbbsnum,b.LastPost,b.readme,b.Board_Setting,b.ParentID,b.Depth,b.ParentStr,b.child,c.* from board b inner join config c on b.sid=c.id where b.boardid="&boardid
	end if
end if
set rs=conn.execute(sql)
if rs.eof and rs.bof then
	response.write "错误的系统参数，很可能您指定了错误的论坛名称，请从有效连接进入。<br>或者您选择了已经被删除的自选风格，请<a href=Ecookies.asp>清空您在本论坛的Cookies信息</a>。"
	response.end
else
	Forum_info=split(rs("Forum_info"),",")
	Forum_Setting=split(rs("Forum_setting"),",")
	Forum_ads=split(rs("Forum_ads"),"$")
	Forum_body=split(rs("Forum_body"),"|||")
	Forum_user=split(rs("Forum_user"),",")

	Forum_userface=split(rs("Forum_userface"),"|")
	Forum_PostFace=split(rs("Forum_PostFace"),"|")
	Forum_emot=split(rs("Forum_emot"),"|")
	Forum_userfaceNum=ubound(Forum_userface)
	Forum_PostFaceNum=ubound(Forum_PostFace)
	Forum_emotNum=ubound(Forum_emot)

	Copyright=rs("Forum_copyright")
	Version="   Powered by：<a href=http://www.dvbbs.net/download.asp>Dvbbs Version 6.0.0</a>"
	badwords=rs("badwords")
	Splitwords=rs("Splitwords")
	StopReadme=rs("StopReadme")
	Forum_pic=split(rs("Forum_pic"),",")
	Forum_boardpic=split(rs("Forum_boardpic"),",")
	Forum_TopicPic=split(rs("Forum_Topicpic"),",")
	Forum_statePic=split(rs("Forum_statepic"),",")
	Forum_upload=rs("Forum_upload")
	Forum_ubb=split(rs("Forum_ubb"),",")
	Forum_sn=rs("Forum_sn")
	MaxOldID=rs("Maxoldid")
	NowUseBbs=rs("NowUseBbs")
	AllPostTable=split(rs("AllPostTable"),"|")
	AllPostTableName=split(rs("AllPostTableName"),"|")

	OpenTime=split(forum_setting(70),"|")
	if Cint(forum_setting(69))=1 and ubound(OpenTime)=1 then
		if IsNumeric(OpenTime(0)) and IsNumeric(OpenTime(1)) then
			if Hour(Now)<Cint(OpenTime(0)) or Hour(Now)>Cint(OpenTime(1)) then
				response.write "本论坛在<B>"&OpenTime(0)&"</B>至<B>"&OpenTime(1)&"</B>点开放，请在规定时间内访问，谢谢。"
				response.end
			end if
		end if
	end if
	if not FoundBoard then
		cookiepath=rs("cookiepath")
		Maxonline=rs("Maxonline")
		skinid=rs("id")
		tid=rs("tid")
		stats="论坛信息"
		Maxonline=rs("Maxonline")
		MaxonlineDate=rs("MaxonlineDate")
		if Clng(forum_setting(26))>0 then
			if Forum_OnlineNum>Clng(forum_setting(26)) then
				if membername="" then
				response.write "当前论坛在线已经超过<B>"&forum_setting(26)&"</B>人，请稍后访问。"
				response.end
				else
				set ors=conn.execute("select username from online where username='"&membername&"'")
				if rs.eof and rs.bof then
				response.write "当前论坛在线已经超过<B>"&forum_setting(26)&"</B>人，请稍后访问。"
				response.end
				end if
				end if
			end if
		end if
	else
		LastPostInfo=split(rs(5),"$")
		Board_Setting=split(rs("board_setting"),",")
		if not isdate(LastPostInfo(2)) then LastPostInfo(2)=Now()
		LastPostTime=LastPostInfo(2)
		boardtype=rs("boardtype")
		lockboard=Board_Setting(0)
		boardmasterlist=rs("boardmaster")
		todaynum=rs(2)
		LastTopicNum=rs("LastTopicNum")
		LastBbsNum=rs("LastBbsNum")
		tid=rs("tid")
		stats=boardtype
		BoardReadme=rs("Readme")
		BoardParentStr=rs("ParentStr")
		BoardParentID=rs("ParentID")
		BoardDepth=rs("depth")
		BoardChild=rs("child")
		skinid=rs("id") 
		if datediff("d",LastPostTime,Now())>0 then
			conn.execute("update board set todaynum=0 where boardid="&boardid)
		end if
		isaudit=Cint(Board_Setting(3))
		OpenTime=split(Board_Setting(22),"|")
		if Cint(Board_Setting(21))=1 and ubound(OpenTime)=1 then
			if IsNumeric(OpenTime(0)) and IsNumeric(OpenTime(1)) then
				if Hour(Now)<Cint(OpenTime(0)) or Hour(Now)>Cint(OpenTime(1)) then
					response.write "本论坛在<B>"&OpenTime(0)&"</B>至<B>"&OpenTime(1)&"</B>点开放，请在规定时间内访问，谢谢。"
					response.end
				end if
			end if
		end if
		if Clng(Board_Setting(18))>0 then
			if Forum_OnlineNum>Clng(Board_Setting(18)) then
				if membername="" then
				response.write "当前论坛在线已经超过<B>"&Board_Setting(18)&"</B>人，请稍后访问。"
				response.end
				else
				set ors=conn.execute("select username from online where username='"&membername&"'")
				if rs.eof and rs.bof then
				response.write "当前论坛在线已经超过<B>"&Board_Setting(18)&"</B>人，请稍后访问。"
				response.end
				end if
				end if
			end if
		end if
	end if
end if
rs.close
set rs=nothing

if cint(Forum_Setting(19))=1 then
	Dim SplitReflashPage
	Dim DoReflashPage
	DoReflashPage=false
	if trim(forum_setting(64))<>"" then
		SplitReflashPage=split(forum_setting(64),"|")
		for i=0 to ubound(SplitReflashPage)
			if instr(scriptname,SplitReflashPage(i))>0 then
				DoReflashPage=true
				exit for
			end if
		next
	end if
	if (not isnull(session("ReflashTime"))) and cint(Forum_Setting(20))>0 and DoReflashPage then
		if DateDiff("s",session("ReflashTime"),Now())<cint(Forum_Setting(20)) then
   		response.write "<META http-equiv=Content-Type content=text/html; charset=gb2312><meta HTTP-EQUIV=REFRESH CONTENT=3>本页面起用了防刷新机制，请不要在"&Forum_Setting(20)&"秒内连续刷新本页面<BR>正在打开页面，请稍后……"
		response.end
		else
		session("ReflashTime")=Now()
		end if
	elseif isnull(session("ReflashTime")) and cint(Forum_Setting(20))>0 and DoReflashPage then
		Session("ReflashTime")=Now()
	end if
end if

if (instr(scriptname,"admin")=0 and instr(scriptname,"login")=0 and instr(scriptname,"chklogin")=0) or master then
	if cint(Forum_Setting(21))=1 then
	Response.write StopReadme
	response.end
	end if

	if LockIP(Request.ServerVariables("REMOTE_ADDR")) then
	response.write "您的IP已经被限制不能访问本论坛，请和管理员联系"&Forum_info(5)&"。"
	response.end
	end if
end if

Server.ScriptTimeOut=Forum_Setting(1)

Rem 用户信息
Dim vipuser,boardmaster,master,superboardmaster
Dim reboardid,reid
Dim LastLogin,myuserep,mysex,myusercp,mymoney,mypower,myArticle,myClass
Dim UserGroupID,GroupSetting
vipuser=false
boardmaster=false
superboardmaster=false
master=false
if userid<>"" and isInteger(userid) then
	sql="select u.userclass,u.userid,u.userpassword,u.lastlogin,u.showRe,u.reAnn,u.userGroupID,g.GroupSetting,u.username,u.userWealth,u.Userep,u.Usercp,u.sex,u.UserPower,u.Article from [user] u inner join UserGroups G on u.UserGroupID=g.UserGroupID where u.userid="&userid
	set rs=conn.execute(sql)
	if rs.eof and rs.bof then
		founduser=false
	else
		if trim(rs(2))=trim(memberword) and lcase(trim(membername))=lcase(trim(rs(8))) then
			founduser=true
			select case rs(6)
			case 8
				vipuser=true
			case 3
				boardmaster=true
			case 2
				superboardmaster=true
			case 1
				master=true
			end select
			userid=rs(1)
			lastlogin=rs(3)
			if rs(4)=1 and rs(5)<>"" then
			reBoardID=split(rs(5),"|")(0)
			reID=split(rs(5),"|")(1)
			end if
			UserGroupID=rs(6)
			myuserep=rs(10)
			mysex=rs(12)
			myusercp=rs(11)
			mymoney=rs(9)
			mypower=rs(13)
			myArticle=rs(14)
			myClass=rs(0)
			GroupSetting=split(rs(7),",")
			if userhidden=2 and DateDiff("s",rs(3),Now())>Clng(Forum_Setting(8))*60 then
			conn.execute("update [user] set UserLastIp='"&Request.ServerVariables("REMOTE_ADDR")&"',LastLogin=Now() where userid="&userid)
			end if
		else
			founduser=false
		end if
	end if
	rs.close
	set rs=nothing
end if
if not founduser then
	founduser=false
	userid=0
	set rs=conn.execute("select GroupSetting from UserGroups where UserGroupID=7")
	GroupSetting=split(rs(0),",")
	UserGroupID=7
	rs.close
	set rs=nothing
end if
if boardid>=0 then GetBoardPermission
'是否继承上级版主，顺带取出上级论坛版面信息
'最多只取向上的10级版面信息
if FoundBoard and BoardParentID>0 then
	Dim FBoardID(9),FBoardName(9),FBoardParentID(9),FBoardMaster
	FBoardmaster=false
	i=0
	set rs=conn.execute("select boardid,boardtype,boardmaster,ParentID from board where boardid in ("&BoardParentStr&") order by orders")
	do while not rs.eof
		FBoardID(i)=rs(0)
		FBoardName(i)=rs(1)
		FBoardParentID(i)=rs(3)
		if Cint(Board_Setting(40))=1 and not FBoardMaster then
			if instr("|"&rs(2)&"|","|"&membername&"|")>0 then
				FBoardMaster=true
			else
				FBoardMaster=false
			end if
		end if
		i=i+1
		if i>9 then exit do
	rs.movenext
	loop
	set rs=nothing
	if master then
		boardmaster=true
	else
		if FBoardMaster then
			boardmaster=true
		else
			if instr("|"&boardmasterlist&"|","|"&membername&"|")>0 then
				boardmaster=true
			else
				boardmaster=false
			end if
		end if
	end if
elseif FoundBoard and FoundUser then
	if master then
		boardmaster=true
	else
		if instr("|"&boardmasterlist&"|","|"&membername&"|")>0 then
		boardmaster=true
		else
		boardmaster=false
		end if
	end if
end if
'endtime=timer()
'response.write "页面执行时间："&FormatNumber((endtime-startime)*1000,3)&"毫秒"
'response.end
%>