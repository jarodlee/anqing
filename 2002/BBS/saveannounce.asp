<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/char_board.asp" -->
<!-- #include file="inc/char_login.asp" -->
<!-- #include file="inc/chkinput.asp" -->
<!-- #include file="inc/ubbcode.asp" -->
<!--#include file="md5.asp"-->
<%
dim announceid
dim UserName
dim userPassword
dim useremail
dim Topic
dim body
dim rootid
dim dateTimeStr
dim ip
dim Expression
dim signflag
dim mailflag
dim usercookies
dim ihaveupfile,upfileinfo,upfilelen

if BoardID="" or not isInteger(BoardID) or BoardID=0 then
	Errmsg=Errmsg+"<br>"+"<li>����İ����������ȷ�����Ǵ���Ч�����ӽ��롣"
	founderr=true
else
	BoardID=clng(BoardID)
end if
IP=Request.ServerVariables("REMOTE_ADDR") 
Expression=Checkstr(Request.Form("Expression"))
Topic=Checkstr(trim(request.Form("subject")))
Body=Checkstr(trim(request.form("Content")))
UserName=Checkstr(trim(request.Form("username")))
signflag=Checkstr(trim(request.Form("signflag")))
mailflag=Checkstr(trim(request.Form("emailflag")))
if memberword=trim(request.Form("passwd")) then
	UserPassWord=Checkstr(trim(request.Form("passwd")))
else
	UserPassWord=md5(Checkstr(trim(request.Form("passwd"))))
end if
if request.form("upfilerename")<>"" then
	ihaveupfile=1
	upfileinfo=replace(request.form("upfilerename"),"'","")
	upfilelen=len(upfileinfo)
	upfileinfo=left(upfileinfo,upfilelen-1)
else
	ihaveupfile=0
end if
if signflag="yes" then
	signflag=1
else
	signflag=0
end if
if mailflag="yes" then
	mailflag=1
else
	mailflag=0
end if
	
dim rndnum,num1
if cint(Board_Setting(30))=1 then
	if not (isnull(session("lastpost")) or boardmaster or master or superboardmaster) then
		if DateDiff("s",session("lastpost"),Now())<cint(Board_Setting(31)) then
   		ErrMsg=ErrMsg+"<Br>"+"<li>����̳���Ʒ�������ʱ��Ϊ"&Board_Setting(31)&"�룬���Ժ��ٷ���"
   		FoundErr=True
		end if
	end if
end if
if chkpost=false then
	ErrMsg=ErrMsg+"<Br>"+"<li>���ύ�����ݲ��Ϸ����벻Ҫ���ⲿ�ύ���ԡ�"
	FoundErr=True
end if
if UserName="" or UserPassWord="" then
	username=membername
	UserPassWord=memberword
end if
if UserName="" then
	ErrMsg=ErrMsg+"<Br>"+"<li>����������������"
	FoundErr=True
end if
if Topic="" then
	FoundErr=True
	ErrMsg=ErrMsg+"<Br>"+"<li>���ӵ����ⲻӦΪ�ա�"
elseif strLength(topic)>50 then
	FoundErr=True
	ErrMsg=ErrMsg+"<Br>"+"<li>�������ⳤ�Ȳ��ܳ���50"
end if
if strLength(body)>Clng(Board_Setting(16)) then
	ErrMsg=ErrMsg+"<Br>"+"<li>�������ݲ��ô���" & CSTR(Board_Setting(16)) & "bytes"
	FoundErr=true
end if
if body="" then
	ErrMsg=ErrMsg+"<Br>"+"<li>û����д���ݡ�"
	FoundErr=true
end if
session("lastpost")=Now()
if founderr then
	call nav()
	call head_var(2,0,"","")
	call dvbbs_error()
else
	stats="�������ӳɹ�"
	call nav()
	call head_var(1,BoardDepth,0,0)
	call main()
	if founderr then call dvbbs_error()
end if
call footer()

sub main()
if Cint(Board_Setting(43))=1 then
	Errmsg=Errmsg+"<br>"+"<li>����̳�Ѿ�������Ա�����˲���������"
	founderr=true
	exit sub
end if
if Cint(GroupSetting(3))=0 then
	Errmsg=Errmsg+"<br>"+"<li>��û���ڱ���̳������Ȩ�ޣ���<a href=login.asp>��½</a>����ͬ����Ա��ϵ��"
	founderr=true
	exit sub
end if
if cint(Board_Setting(2))=1 then
	if not founduser then
		Errmsg=Errmsg+"<br>"+"<li>����̳Ϊ��֤��̳����<a href=login.asp>��½</a>��ȷ�������û����Ѿ��õ�����Ա����֤����롣"
		founderr=true
		exit sub
	else
		if chkboardlogin(boardid,membername)=false then
		Errmsg=Errmsg+"<br>"+"<li>����̳Ϊ��֤��̳����ȷ�������û����Ѿ��õ�����Ա����֤����롣"
		founderr=true
		exit sub
		end if
	end if
end if

usercookies=request.Cookies("aspsky")("usercookies")
if isnull(usercookies) or usercookies="" then usercookies=3

if chkuserlogin(username,userpassword,usercookies,2)=false then
	errmsg=errmsg+"<br>"+"<li>�����û����������ڣ���������������󣬻��������ʺ��ѱ�����Ա������"
	founderr=true
	exit sub
end if

if cint(Board_Setting(1))=1 then
	if Cint(GroupSetting(37))=0 then
		Errmsg=ErrMsg+"<Br>"+"<li>��û��Ȩ�޽���������̳��"
		founderr=true
		exit sub
	end if
end if
rem ����������Ϣ
dim locktopic
if master or superboardmaster or boardmaster then
	isaudit=0
	locktopic=0
elseif isaudit=1 and not (master or superboardmaster or boardmaster) then
	isaudit=1
	locktopic=3
else
	isaudit=0
	locktopic=0
end if
dim LastPost,LastPost_1,uploadpic_n,Forumupload,u
dim LastPostTimes
DateTimeStr=replace(replace(CSTR(NOW()+Forum_Setting(0)/24),"����",""),"����","")
'���������
sql="insert into topic (Title,Boardid,PostUsername,PostUserid,DateAndTime,Expression,LastPost,LastPostTime,PostTable,locktopic) values ('"&topic&"',"&boardid&",'"&username&"',"&userid&",'"&DateTimeStr&"','"&Expression&"','$$"&DateTimeStr&"$$$$','"&DateTimeStr&"','"&NowUseBbs&"',"&locktopic&")"
conn.execute(sql)

set rs=conn.execute("select top 1 topicid from topic order by topicid desc")
rootid=rs(0)
Forum_upload="gif,jpg,jpeg,bmp,zip,rar,html,swf,mid,midi,flash,rm,ra,asf,avi,wmv,exe,xls,dos,ftp,mp3,m3u,txt,mdb,dll,IMG"
Forumupload=split(Forum_upload,",")
for u=0 to ubound(Forumupload)
	if instr(body,"[upload="&Forumupload(u)&"]") or instr(body,"."&Forumupload(u)&"") or instr(body,"["&Forumupload(u)&"]")  then
		uploadpic_n=Forumupload(u)
		exit for
	end if
next
if instr(body,"viewfile.asp?ID=") then uploadpic_n="down"

'����ظ���
Sql="insert into "&NowUseBbs&"(Boardid,ParentID,username,topic,body,DateAndTime,length,rootid,layer,orders,ip,Expression,locktopic,signflag,emailflag,isbest,PostUserID,isupload,isaudit) values ("&boardid&",0,'"&username&"','"&topic&"','"&body&"','"&DateTimeStr&"','"&strlength(body)&"',"&rootid&",1,0,'"&ip&"','"&Expression&"',"&locktopic&","&signflag&","&mailflag&",0,"&userid&","&ihaveupfile&","&isaudit&")"
conn.execute(sql)

set rs=conn.execute("select top 1 Announceid from "&NowUseBbs&" order by Announceid desc")
Announceid=rs(0)

if ihaveupfile=1 then conn.execute("update dv_upfile set F_AnnounceID='"&rootid&"|"&AnnounceID&"',F_Readme='"&Topic&"' where F_ID in ("&upfileinfo&")")

LastPost=replace(username,"$","") & "$" & Announceid & "$" & DateTimeStr & "$" & replace(cutStr(body,20),"$","") & "$" & uploadpic_n & "$" & UserID & "$" & rootid & "$" & BoardID
LastPost=reubbcode(replace(LastPost,"'",""))

conn.execute("update topic set LastPost='"&LastPost&"' where topicid="&rootid)

LastPost_1=replace(username,"$","") & "$" & Announceid & "$" & DateTimeStr & "$" & replace(cutStr(Topic,20),"$","") & "$" & uploadpic_n & "$" & UserID & "$" & rootid & "$" & BoardID
LastPost_1=reubbcode(replace(LastPost_1,"'",""))

if isaudit=0 then
Dim UpdateBoardID
UpdateBoardID=BoardParentStr & "," & BoardID
if datediff("d",LastPostTime,Now())=0 then
	sql="update board set lastbbsnum=lastbbsnum+1,lasttopicnum=lasttopicnum+1,todaynum=todaynum+1,LastPost='"&LastPost_1&"' where boardid in ("&UpdateBoardID&")"
else
	sql="update board set lastbbsnum=lastbbsnum+1,lasttopicnum=lasttopicnum+1,todaynum=1,LastPost='"&LastPost_1&"' where  boardid in ("&UpdateBoardID&")"
end if
conn.execute(sql)

Dim updateinfo
set rs=conn.execute("select LastPost,TodayNum,MaxPostNum from config where active=1")
LastPostTimes=split(rs(0),"$")
LastPostTime=LastPostTimes(2)
if not isdate(LastPostTime) then LastPostTime=Now()
if datediff("d",LastPostTime,Now())=0 then
	if rs(1)+1>rs(2) then updateinfo=",MaxPostNum=todaynum+1"
	sql="update config set topicnum=topicnum+1,bbsnum=bbsnum+1,todayNum=todayNum+1,LastPost='"&LastPost&"' "&updateinfo&" where active=1"
else
	sql="update config set topicnum=topicnum+1,bbsnum=bbsnum+1,yesterdaynum="&rs(1)&",todayNum=1,LastPost='"&LastPost&"' where active=1"
end if
conn.execute(sql)
end if

dim PostRetrunName
select case Board_Setting(17)
case 1
	response.write "<meta http-equiv=refresh content=""3;URL=index.asp"">"
	PostRetrunName="��ҳ"
case 2
	response.write "<meta http-equiv=refresh content=""3;URL=list.asp?boardid="&boardid&""">"
	PostRetrunName="������������̳"
case 3
	if isaudit=1 then
	response.write "<meta http-equiv=refresh content=""3;URL=list.asp?boardid="&boardid&""">"
	PostRetrunName="�����������ӱ��뾭����Ա��˺󷽿ɼ�"
	else
	response.write "<meta http-equiv=refresh content=""3;URL=dispbbs.asp?boardid="&boardid&"&id="&rootid&""">"
	PostRetrunName="�������������"
	end if
end select
set rs=nothing
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr align=center><th width="100%">״̬��<%=stats%></td>
</tr><tr><td width="100%" class=tablebody1>
��ҳ�潫��3����Զ�����<%=PostRetrunName%>��<b>������ѡ�����²�����</b><br><ul>
<li><a href="index.asp">������ҳ</a></li>
<li><a href="list.asp?boardid=<%=boardid%>"><%=boardtype%></a></li>
<li><%if isaudit=1 then%><%=PostRetrunName%><%else%><a href="dispbbs.asp?boardid=<%=boardid%>&id=<%=rootid%>&star=<%=request("star")%>#<%=announceid%>"><%=PostRetrunName%></a><%end if%></li>
</ul></td></tr></table>
<%
end sub
%>