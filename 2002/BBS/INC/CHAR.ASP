<%
Rem ==========论坛通用函数=========
'页面错误提示信息
sub dvbbs_error()
%>
<br>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1 style="width:75%">
<tr align=center>
<th width="100%" height=25 colspan=2>论坛错误信息
</td>
</tr>
<tr>
<td width="100%" class=tablebody1 colspan=2>
<b>产生错误的可能原因：</b><br><br>
<li>您是否仔细阅读了<a href="boardhelp.asp?boardid=<%=boardid%>">帮助文件</a>，可能您还没有登陆或者不具有使用当前功能的权限。
<%=errmsg%>
</td></tr>
<%if not founduser then%>
<form action="login.asp?action=chk" method=post>
    <tr>
    <th valign=middle colspan=2 align=center height=25>请输入您的用户名、密码登陆</td></tr>
    <tr>
    <td valign=middle class=tablebody1>请输入您的用户名</td>
    <td valign=middle class=tablebody1><INPUT name=username type=text> &nbsp; <a href=reg.asp>没有注册？</a></td></tr>
    <tr>
    <td valign=middle class=tablebody1>请输入您的密码</font></td>
    <td valign=middle class=tablebody1><INPUT name=password type=password> &nbsp; <a href=lostpass.asp>忘记密码？</a></td></tr>
    <tr>
    <td class=tablebody1 valign=top width=30% ><b>Cookie 选项</b><BR> 请选择你的 Cookie 保存时间，下次访问可以方便输入。</td>
    <td valign=middle class=tablebody1>                <input type=radio name=CookieDate value=0 checked>不保存，关闭浏览器就失效<br>
                <input type=radio name=CookieDate value=1>保存一天<br>
                <input type=radio name=CookieDate value=2>保存一月<br>
                <input type=radio name=CookieDate value=3>保存一年<br>                </td></tr>
    <tr>
    <td valign=top width=30% class=tablebody1><b>隐身登陆</b><BR> 您可以选择隐身登陆，论坛会员将在用户列表看不到您的信息。</td>
    <td valign=middle class=tablebody1>                <input type=radio name=userhidden value=2 checked>正常登陆<br>
                <input type=radio name=userhidden value=1>隐身登陆<br>
                </td></tr>
	<input type=hidden name=comeurl value="<%=Request.ServerVariables("HTTP_REFERER")%>">
    <tr>
    <td class=tablebody2 valign=middle colspan=2 align=center><input type=submit name=submit value="登 陆">&nbsp;&nbsp;<input type=button name="back" value="返 回" onclick="location.href='<%=Request.ServerVariables("HTTP_REFERER")%>'"></td></tr>
</form>
<%else%>
    <tr>
    <td class=tablebody2 valign=middle colspan=2 align=center><a href="<%=Request.ServerVariables("HTTP_REFERER")%>"><<返回上一页</a></td></tr>
<%end if%>
</table>
<%
end sub

sub dvbbs_suc()
%>
<br>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr align=center>
<th width="100%">论坛成功信息
</td>
</tr>
<tr>
<td width="100%" class=tablebody1>
<b>操作成功：</b><br><br>
<%=sucmsg%>
</td></tr>
<tr align=center><td width="100%" class=tablebody2>
<a href="<%=Request.ServerVariables("HTTP_REFERER")%>"> << 返回上一页</a>
</td></tr>
</table>
<%
end sub

rem 过滤字符
function ChkBadWords(fString)
    dim bwords,ii
    if not(isnull(BadWords) or isnull(fString)) then
    bwords = split(BadWords, "|")
    for ii = 0 to ubound(bwords)
        fString = Replace(fString, bwords(ii), string(len(bwords(ii)),"*")) 
    next
    ChkBadWords = fString
    end if
end function

Rem 过滤HTML代码
function HTMLEncode(fString)
if not isnull(fString) then
    fString = replace(fString, ">", "&gt;")
    fString = replace(fString, "<", "&lt;")

    fString = Replace(fString, CHR(32), "&nbsp;")
    fString = Replace(fString, CHR(9), "&nbsp;")
    fString = Replace(fString, CHR(34), "&quot;")
    fString = Replace(fString, CHR(39), "&#39;")
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
    fString = Replace(fString, CHR(10), "<BR> ")

    fString=ChkBadWords(fString)
    HTMLEncode = fString
end if
end function

Rem 过滤表单字符
function HTMLcode(fString)
if not isnull(fString) then
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10) & CHR(10), "</P><P>")
    fString = Replace(fString, CHR(10), "<BR>")
    HTMLcode = fString
end if
end function

'用户IP限制
function LockIP(sip)
	dim str1,str2,str3,str4
	dim num
	LockIP=false
	if isnumeric(left(sip,2)) then
		str1=left(sip,instr(sip,".")-1)
		sip=mid(sip,instr(sip,".")+1)
		str2=left(sip,instr(sip,".")-1)
		sip=mid(sip,instr(sip,".")+1)
		str3=left(sip,instr(sip,".")-1)
		str4=mid(sip,instr(sip,".")+1)
		if isNumeric(str1)=0 or isNumeric(str2)=0 or isNumeric(str3)=0 or isNumeric(str4)=0 then
	
		else
			num=cint(str1)*256*256*256+cint(str2)*256*256+cint(str3)*256+cint(str4)-1
			sql="select count(*) from LockIP where ip1 <="&num&" and ip2 >="&num&""
			set rs=conn.execute(sql)
			if rs(0)>0 then 
				LockIP=true
			end if
			set rs=nothing
		end if
	end if
end function

Rem 判断发言是否来自外部
function ChkPost()
	dim server_v1,server_v2
	chkpost=false
	server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
	server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
	if mid(server_v1,8,len(server_v2))<>server_v2 then
		chkpost=false
	else
		chkpost=true
	end if
end function

Rem 判断用户来源
function address(sip)
	dim str1,str2,str3,str4
	dim num
	dim country,city
	dim irs
	if isnumeric(left(sip,2)) then
	if sip="127.0.0.1" then sip="192.168.0.1"
	str1=left(sip,instr(sip,".")-1)
	sip=mid(sip,instr(sip,".")+1)
	str2=left(sip,instr(sip,".")-1)
	sip=mid(sip,instr(sip,".")+1)
	str3=left(sip,instr(sip,".")-1)
	str4=mid(sip,instr(sip,".")+1)
	if isNumeric(str1)=0 or isNumeric(str2)=0 or isNumeric(str3)=0 or isNumeric(str4)=0 then

	else
		num=cint(str1)*256*256*256+cint(str2)*256*256+cint(str3)*256+cint(str4)-1
		sql="select Top 1 country,city from address where ip1 <="&num&" and ip2 >="&num&""
		set irs=server.createobject("adodb.recordset")
		irs.open sql,conn,1,1
		if irs.eof and irs.bof then 
		country="亚洲"
		city=""
		else
		country=irs(0)
		city=irs(1)
		end if
		irs.close
		set irs=nothing
	end if
	address=country&city
	else
	address="未知"
	end if
end function

function iif(expression,returntrue,returnfalse)
	if expression=0 then
		iif=returnfalse
	else
		iif=returntrue
	end if
end function

function iiif(express,expression,returntrue,returnfalse)
	if express>expression then
		iiif=returnfalse
	else
		iiif=returntrue
	end if
end function

function iimg(expression,returnfalse,returntrue)
	if expression="" or isnull(expression) then
		iimg=returnfalse
	else
		iimg=returntrue
	end if
end function

Rem 过滤SQL非法字符
function checkStr(str)
	if isnull(str) then
		checkStr = ""
		exit function 
	end if
	checkStr=replace(str,"'","''")
end function

Rem 用户在线
sub activeonline()
dim ComeFrom,actCome,statuserid
statuserid=replace(Request.ServerVariables("REMOTE_HOST"),".","")
if not founduser then
	session("userid")=statuserid
	sql="select id,boardid from online where id="&cstr(session("userid"))
	set rs=conn.execute(sql)
	if rs.eof and rs.bof then
		ComeFrom=""
		actCome=""
		sql="insert into online(id,username,userclass,ip,startime,lastimebk,boardid,browser,stats,actforip,UserGroupID,actCome,userhidden) values ("&statuserid&",'客人','客人','"&Request.ServerVariables("REMOTE_HOST")&"',Now(),Now(),"&boardid&",'"&Request.ServerVariables("HTTP_USER_AGENT")&"','"&replace(stats,"'","")&"','"&Request.ServerVariables("HTTP_X_FORWARDED_FOR")&"',7,'"&actCome&"',"&userhidden&")"
	else
		sql="update online set lastimebk=Now(),boardid="&boardid&",stats='"&replace(stats,"'","")&"' where id="&cstr(session("userid"))
	end if
	conn.execute(sql)
else
	if founderr then
	boardid=0
	stats="错误信息"
	end if
	sql="select id,boardid from online where userid="&userid
	set rs=conn.execute(sql)
	if rs.eof and rs.bof then
	ComeFrom=""
	actCome=""
		sql="insert into online(id,username,userclass,ip,startime,lastimebk,boardid,browser,stats,actforip,UserGroupID,actCome,userhidden,userid) values ("&statuserid&",'"&membername&"','"&memberclass&"','"&Request.ServerVariables("REMOTE_HOST")&"',Now(),Now(),"&boardid&",'"&Request.ServerVariables("HTTP_USER_AGENT")&"','"&replace(stats,"'","")&"','"&Request.ServerVariables("HTTP_X_FORWARDED_FOR")&"',"&UserGroupID&",'"&actCome&"',"&userhidden&","&userid&")"
	else
		sql="update online set lastimebk=Now(),boardid="&boardid&",stats='"&replace(stats,"'","")&"' where userid="&userid
	end if
	conn.execute(sql)
	rs.close
	if session("userid")<>"" then
	Conn.Execute("delete from online where id="&session("userid"))
	session("userid")=""
	end if
end if
set rs=nothing
Rem 删除超时用户
sql="Delete FROM online WHERE DATEDIFF('s', lastimebk, now()) > "&Forum_Setting(8)&"*60"
Conn.Execute sql
end sub

sub footer()
endtime=timer()
%>
<p>
<TABLE cellSpacing=0 cellPadding=0 width="<%=Forum_body(12)%>" border=0 align=center>
<tr><td align=center>
<%=Forum_ads(1)%>
</td></tr>
<tr><td align=center>
<%=Version%><br>
<%=Copyright%><%if Cint(forum_setting(30))=1 then%> , 页面执行时间：<%=FormatNumber((endtime-startime)*1000,3)%>毫秒<%end if%>
</td></tr>
</table>
</body>
</html>
<%
	CloseDatabase
	set myCache=nothing
end sub

sub nav()
%>
<html>
<head>
<META http-equiv=Content-Type content=text/html; charset=gb2312>
<meta name=keywords content="动网先锋,动网论坛,dvbbs">
<title><%=Forum_info(0)%>--<%=HTMLEncode(stats)%></title>
<!--#include file="Forum_css.asp"-->
<!--#include file="Forum_js.asp"-->
</head>
<body <%=Forum_body(11)%> onmousemove="HideMenu()">
<div id=menuDiv style='Z-INDEX: 2; VISIBILITY: hidden; WIDTH: 1px; POSITION: absolute; HEIGHT: 1px; BACKGROUND-COLOR: #9cc5f8'></div>
<table cellspacing=0 cellpadding=0 align=center style="border:1px <%=Forum_body(27)%> solid; border-top-width: 0px; border-right-width: 1px; border-bottom-width: 0px; border-left-width: 1px;width:<%=Forum_body(12)%>;">
<tr><td width=100%>
<table width=100% align=center border=0 cellspacing=0 cellpadding=0>
<tr><td class=TopDarkNav height=9></td></tr>
<tr><td height=70 class=TopLighNav2>
<TABLE border=0 width="100%" align=center>
<TR>
<TD align=left width="25%"><a href="<%= Forum_info(3) %>"><img border=0 src='<%= Forum_info(6) %>'></a></TD>
<TD Align=center width="65%">
<%if isnull(Forum_ads(0)) or Forum_ads(0)="" then%>
<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="468" height="60"><param name=movie value="http://www.dvbbs.net/skin/default/dvbanner.swf"><param name=quality value=high><param name=menu value=false><embed src="http://www.dvbbs.net/skin/default/dvbanner.swf" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="468" height="60"></embed></object>
<%else%>
<%=Forum_ads(0)%>
<%end if%>
</td>
<td align=right style="line-height: 15pt" width="10%">
<a href=#><span style="CURSOR: hand" onClick="window.external.AddFavorite('<%=Forum_info(1)%>', '<%=Forum_info(0)%>')">加入收藏</span></a>
<br><a href="mailto:<%=Forum_info(5)%>">联系我们</a>
<br><a href="boardhelp.asp?boardid=<%=boardid%>">论坛帮助</a>
</td>
</td></tr>
</table>
</td></tr>
<tr><td class=TopLighNav height=9></td></tr>
        <tr> 
          <td class=TopLighNav1 height=22  valign="middle">&nbsp;&nbsp;
<%if not founduser then%>
<a href="login.asp">登陆</a> <img src=<%=Forum_info(7)%>navspacer.gif align=absmiddle> <a href="reg.asp">注册</a>
<%else%>
<%if userhidden=2 then%><a href="cookies.asp?action=hidden&userid=<%=userid%>">隐身</a><%else%><a href="cookies.asp?action=online&userid=<%=userid%>">上线</a><%end if%> <img src=<%=Forum_info(7)%>navspacer.gif align=absmiddle> <a href="login.asp">重登陆</a> <img src=<%=Forum_info(7)%>navspacer.gif align=absmiddle> <a href="usermanager.asp" onMouseOver='ShowMenu(manage,100)'>用户控制面板</a>
<%end if%>
 <img src=<%=Forum_info(7)%>navspacer.gif align=absmiddle> <a href="query.asp?boardid=<%=boardid%>">搜索</a> <img src=<%=Forum_info(7)%>navspacer.gif align=absmiddle> <a href="#" onMouseOver='ShowMenu(stylelist,100)'>自选风格</a> <img src=<%=Forum_info(7)%>navspacer.gif align=absmiddle> <a href="boardstat.asp?boardid=<%=boardid%>" onMouseOver='ShowMenu(boardstat,100)'>论坛状态</a> <img src=<%=Forum_info(7)%>navspacer.gif align=absmiddle> <a href="show.asp?boardid=<%=boardid%>" onMouseOver='ShowMenu(downlist,100)'>论坛展区</a> <%if founduser then%> <img src=<%=Forum_info(7)%>navspacer.gif align=absmiddle> <a href="logout.asp">退出</a><%else%> <img src=<%=Forum_info(7)%>navspacer.gif align=absmiddle> <a href="dispuser.asp?boardid=<%=boardid%>&action=permission">我能做什么</a><%end if%>
<%if master then response.write " <img src="&Forum_info(7)&"navspacer.gif align=absmiddle> <a href=admin_index.asp>管理</a> <img src="&Forum_info(7)&"navspacer.gif align=absmiddle> <a href=""recycle.asp"">回收站</a>"%>
			</td>
        </tr>
</table>
</td></tr>
</table>
<%
if Cint(GroupSetting(0))=0 and (instr(scriptname,"reg.asp")=0 and instr(scriptname,"login.asp")=0) then
	Errmsg=Errmsg+"<br>"+"<li>您没有浏览本论坛的权限，请<a href=login.asp>登陆</a>或者同管理员联系。"
	call head_var(2,0,"","")
	call dvbbs_error()
	call footer()
	response.end
end if
end sub

'入口参数
'IsBoard=1论坛分版面导航，IsBoard=0论坛其他页面导航，GetTitle论坛其他页面上级页面，GetUrl论坛其他页面上级页面URL
'Depth论坛分版面导航中论坛深度，其他页面设置为0
sub head_var(IsBoard,idepth,GetTitle,GetUrl)
%>
<table cellspacing=1 cellpadding=3 align=center border=0 width="<%=Forum_body(12)%>">
<tr>
<%if not founduser then%>
<td height=25>
<BR>
>> <%if foundboard then%><%=BoardReadme%><%else%>欢迎光临 <B><%=Forum_info(0)%></B><%end if%>
<%else%>
<td width=65% >
</td><td width=35% align=right>
<%if Cint(newincept())>Cint(0) then%>
<bgsound src="<%=Forum_info(7)&Forum_statePic(8)%>" border=0>
<%if Cint(forum_setting(10))=1 then%>
<script language=JavaScript>openScript('messanger.asp?action=read&id=<%=inceptid(1)%>&sender=<%=inceptid(2)%>',500,400)</script>
<%end if%>
<img src=<%=Forum_info(7)&Forum_boardpic(9)%>> <a href="usersms.asp?action=inbox">我的收件箱</a> (<a href="javascript:openScript('messanger.asp?action=read&id=<%=inceptid(1)%>&sender=<%=inceptid(2)%>',500,400)"><font color="<%=Forum_body(8)%>"><%=newincept()%> 新</font></a>)
<%else%>
<img src=<%=Forum_info(7)&Forum_boardpic(8)%>> <a href="usersms.asp?action=inbox">我的收件箱</a> (<font color=gray>0 新</font>)
<%end if%>
<%end if%>
</td></tr>
</table>
<table cellspacing=1 cellpadding=3 align=center class=tableBorder2>
<tr><td height=25 valign=middle>
<img src="<%=Forum_info(7)&Forum_pic(12)%>" align=absmiddle> <a href=index.asp><%=Forum_info(0)%></a> → 
<%
if IsBoard=1 then
	if BoardParentID>0 then
	for i=0 to idepth-1
		response.write "<a href=list.asp?boardid="&FBoardID(i)&">"&FBoardName(i)&"</a> →  "
		if i>9 then exit for
	next
	end if
	if request("CatLog")="NN" then
	Response.Cookies("BoardList")(BoardID & "BoardID")= "NNotShow"
	end if
	response.write "<a href=list.asp?boardid="&boardid&">"&boardtype&"</a> →  " & HTMLEncode(stats)
	if request.cookies("BoardList")(boardid & "BoardID")="NNotShow" then
	response.write "&nbsp;<a href=""?BoardID="&boardid&"&cBoardid="&boardid&"&Catlog=Y"" title=""展开论坛列表"">[展开]</a>"
	end if
elseif IsBoard=2 then
	response.write HTMLEncode(stats)
else
	response.write "<a href="&GetUrl&">"&GetTitle&"</a> → " & HTMLEncode(stats)
end if
%>
<a name=top></a>
</td></td>
</table>
<br>
<%
end sub

'统计留言
function newincept()
rs=conn.execute("Select Count(id) From Message Where flag=0 and issend=1 and delR=0 And incept='"& membername &"'")
    newincept=rs(0)
	set rs=nothing
	if isnull(newincept) then newincept=0
end function
function inceptid(stype)
	set rs=conn.execute("Select top 1 id,sender From Message Where flag=0 and issend=1 and delR=0 And incept='"& membername &"'")
	if stype=1 then
	inceptid=rs(0)
	else
	inceptid=rs(1)
	end if
	set rs=nothing
end function

Rem 获得版面用户组权限配置
function GetBoardPermission()
	dim pmrs
	GetBoardPermission=false
	set pmrs=conn.execute("select PSetting from BoardPermission where Boardid="&Boardid&" and GroupID="&UserGroupID)
	if not (pmrs.eof and pmrs.bof) then
		GetBoardPermission=true
		GroupSetting=split(pmrs(0),",")
	else
		GetBoardPermission=false
	end if
	if FoundUser then
	set pmrs=conn.execute("select uc_Setting from UserAccess where uc_BoardID="&BoardID&" and uc_UserID="&userID)
	if not(pmrs.eof and pmrs.bof) then
		UserPermission=split(pmrs(0),",")
		GroupSetting=split(pmrs(0),",")
		FoundUserPer=true
	end if
	end if
	set pmrs=nothing
end function


Rem 判断数字是否整形
function isInteger(para)
       on error resume next
       dim str
       dim l,i
       if isNUll(para) then 
          isInteger=false
          exit function
       end if
       str=cstr(para)
       if trim(str)="" then
          isInteger=false
          exit function
       end if
       l=len(str)
       for i=1 to l
           if mid(str,i,1)>"9" or mid(str,i,1)<"0" then
              isInteger=false 
              exit function
           end if
       next
       isInteger=true
       if err.number<>0 then err.clear
end function

function allonline()
dim tmprs
tmprs=conn.execute("Select count(*) from online") 
allonline=tmprs(0) 
set tmprs=nothing 
if isnull(allonline) then allonline=0
end function

'**********************************************
'	vbs Cache类
'
'	属性valid，是否可用，取值前判断
'	属性name，cache名，新建对象后赋值
'	方法add(值,到期时间)，设置cache内容
'	属性value，返回cache内容
'	属性blempty，是否未设置值
'	方法makeEmpty，释放内存，测试用
'	方法equal(变量1)，判断cache值是否和变量1相同
'	方法expires(time)，修改过期时间为time
'	木鸟  2002.12.24
'	http://www.aspsky.net/
'**********************************************

class Cache
	private obj			'cache内容
	private expireTime		'过期时间
	private expireTimeName	'过期时间application名
	private cacheName		'cache内容application名
	private path			'uri
	
	private sub class_initialize()
		path=request.servervariables("url")
		path=left(path,instrRev(path,"/"))
	end sub
	
	private sub class_terminate()
	end sub
	
	public property get blEmpty
		'是否为空
		if isempty(obj) then
			blEmpty=true
		else
			blEmpty=false
		end if
	end property
	
	public property get valid
		'是否可用(过期)
		if isempty(obj) or not isDate(expireTime) then
			valid=false
		elseif CDate(expireTime)<now then
				valid=false
		else
			valid=true
		end if
	end property
	
	public property let name(str)
		'设置cache名
		cacheName=str & path
		obj=application(cacheName)
		expireTimeName=str & "expires" & path
		expireTime=application(expireTimeName)
	end property
	
	public property let expires(tm)
		'重设置过期时间
		expireTime=tm
		application.lock
		application(expireTimeName)=expireTime
		application.unlock
	end property
	
	public sub add(var,expire)
		'赋值
		if isempty(var) or not isDate(expire) then
			exit sub
		end if
		obj=var
		expireTime=expire
		application.lock
		application(cacheName)=obj
		application(expireTimeName)=expireTime
		application.unlock
	end sub
	
	public property get value
		'取值
		if isempty(obj) or not isDate(expireTime) then
			value=null
		elseif CDate(expireTime)<now then
			value=null
		else
			value=obj
		end if
	end property
	
	public sub makeEmpty()
		'释放application
		application.lock
		application(cacheName)=empty
		application(expireTimeName)=empty
		application.unlock
		obj=empty
		expireTime=empty
	end sub
	
	public function equal(var2)
		'比较
		if typename(obj)<>typename(var2) then
			equal=false
		elseif typename(obj)="Object" then
			if obj is var2 then
				equal=true
			else
				equal=false
			end if
		elseif typename(obj)="Variant()" then
			if join(obj,"^")=join(var2,"^") then
				equal=true
			else
				equal=false
			end if
		else
			if obj=var2 then
				equal=true
			else
				equal=false
			end if
		end if
	end function
end class
%>