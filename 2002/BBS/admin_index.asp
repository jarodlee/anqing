<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="md5.asp" -->
<%
dim username
dim password
dim ip
stats=Forum_info(0)&"--控制面板"
select case request("action")
case "admin_left"
	call admin_left()
case "admin_login"
	call admin_login()
case "admin_main"
	call admin_main()
case "admin_head"
	call admin_head()
case else
	call main()
end select
sub main()
if not master or session("flag")="" then
	call admin_login()
else
%>
<html>
<head>
<title><%=Forum_info(0)%>--控制面板</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<frameset cols="180,*" frameborder="NO" border="0" framespacing="0" rows="*"> 
  <frame name="leftFrame" scrolling="AUTO" noresize src="admin_index.asp?action=admin_left" marginwidth="0" marginheight="0">

<frameset rows="20,*"  framespacing="0" border="0" frameborder="0" frameborder="no" border="0">
<frame src="admin_index.asp?action=admin_head" name="head" scrolling="NO" NORESIZE frameborder="0" marginwidth="10" marginheight="0" border="no">
<%if not master or session("flag")="" then%>
  <frame name="main" src="admin_index.asp?action=admin_login" scrolling="AUTO" NORESIZE frameborder="0" marginwidth="10" marginheight="10" border="no">
<%else%>
  <frame name="main" src="admin_index.asp?action=admin_main" scrolling="AUTO" NORESIZE frameborder="0" marginwidth="10" marginheight="10" border="no">
<%end if%>
</frameset>
</frameset>
<noframes>

</body></noframes>
</html>
<%
end if
end sub

sub admin_left()
%>
<title><%=Forum_info(0)%>--管理页面</title>
<style type=text/css>
body  { background:#799AE1; margin:0px; font:9pt 宋体; }
table  { border:0px; }
td  { font:normal 12px 宋体; }
img  { vertical-align:bottom; border:0px; }
a  { font:normal 12px 宋体; color:#000000; text-decoration:none; }
a:hover  { color:#428EFF;text-decoration:underline; }
.sec_menu  { border-left:1px solid white; border-right:1px solid white; border-bottom:1px solid white; overflow:hidden; background:#D6DFF7; }
.menu_title  { }
.menu_title span  { position:relative; top:2px; left:8px; color:#215DC6; font-weight:bold; }
.menu_title2  { }
.menu_title2 span  { position:relative; top:2px; left:8px; color:#428EFF; font-weight:bold; }
</style>
<SCRIPT language=javascript1.2>
function showsubmenu(sid)
{
whichEl = eval("submenu" + sid);
if (whichEl.style.display == "none")
{
eval("submenu" + sid + ".style.display=\"\";");
}
else
{
eval("submenu" + sid + ".style.display=\"none\";");
}
}
</SCRIPT>
<%


REM 管理栏目设置
dim menu(7,10)
menu(0,0)="常规设置"
menu(0,1)="<a href=admin_setting.asp target=main>常规设置信息</a>"
menu(0,2)="<a href=admin_ads.asp target=main>论坛广告设置</a>"
menu(0,3)="<a href=bbseven.asp?action=batch target=_blank>论坛日志管理</a>"
menu(0,4)="<a href=admin_inputfile.asp target=main>初始信息设置</a>"

menu(1,0)="论坛管理"
menu(1,1)="<a href=admin_board.asp?action=add target=main>论坛版面添加</a> | <a href=admin_board.asp target=main>管理</a>"
menu(1,2)="<a href=admin_board.asp?action=permission target=main>论坛权限管理</a>"
menu(1,3)="<a href=admin_boardunite.asp target=main>合并论坛数据</a>"
menu(1,4)="<a href=admin_update.asp target=main>更新论坛数据</a>"
menu(1,5)="<a href=admin_link.asp?action=add target=main>联盟论坛添加</a> | <a href=admin_link.asp target=main>管理</a>"
menu(1,6)="<a href=announcements.asp?boardid=0&action=AddAnn target=_blank>发布论坛公告 | 管理</a>"
menu(1,7)="<a href=admin_menpai.asp target=main>论坛门派管理</a>"

menu(2,0)="用户管理"
menu(2,1)="<a href=admin_user.asp target=main>用户信息管理</a>"
menu(2,2)="<a href=admin_grade.asp?action=add target=main>论坛等级添加</a> | <a href=admin_grade.asp target=main>管理</a>"
menu(2,3)="<a href=admin_wealth.asp target=main>用户积分设置</a>"
menu(2,4)="<a href=admin_message.asp target=main>论坛短信管理</a>"
menu(2,5)="<a href=admin_group.asp?action=addgroup target=main>用户组添加</a> | <a href=admin_group.asp target=main>管理</a>"
menu(2,6)="<a href=admin_admin.asp?action=add target=main>管理员添加</a> | <a href=admin_admin.asp target=main>管理</a>"
menu(2,7)="<a href=admin_mailist.asp target=main>邮件列表</a> | <a href=admin_mailout.asp target=main>列表导出</a>"
menu(2,8)="<a href=admin_update.asp?action=updateuser target=main>更新用户数据</a>"

menu(3,0)="主题和帖子设置"
menu(3,1)="<a href=admin_alldel.asp target=main>批量删除</a>"
menu(3,2)="<a href=admin_alldel.asp?action=moveinfo target=main>批量移动</a>"
menu(3,3)="<a href=recycle.asp target=_blank>回收站管理</a>"
menu(3,4)="<a href=admin_postdata.asp?action=Nowused target=main>当前帖子数据表管理</a>"
menu(3,5)="<a href=admin_postdata.asp target=main>数据表间帖子转换</a>"

menu(4,0)="外观设置"
menu(4,1)="<a href=admin_color.asp target=main>论坛风格CSS设置</a>"
menu(4,2)="<a href=admin_pic.asp target=main>基本图片设置</a>"
menu(4,3)="<a href=admin_skin.asp?action=news target=main>设置模板添加</a> | <a href=admin_skin.asp target=main>管理</a>"
menu(4,4)="<a href=admin_loadskin.asp target=main>CSS风格导出</a> | <a href=admin_loadskin.asp?action=load target=main>导入</a>"

menu(5,0)="替换/限制处理"
menu(5,1)="<a href=admin_badword.asp?reaction=badword target=main>帖子过滤字符</a>"
menu(5,2)="<a href=admin_badword.asp?reaction=splitreg target=main>注册过滤字符</a>"
menu(5,3)="<a href=admin_lockip.asp?action=add target=main>IP来访限定添加</a> | <a href=admin_lockip.asp target=main>管理</a>"
menu(5,4)="<a href=admin_address.asp?action=add target=main>论坛IP库添加</a> | <a href=admin_address.asp target=main>管理</a>"
menu(5,5)="<a href=admin_address.asp?action=upip target=main>导入IP库</a>"

menu(6,0)="数据处理(Access)"
menu(6,1)="<a href=admin_data.asp?action=CompressData target=main>压缩数据库</a>"
menu(6,2)="<a href=admin_data.asp?action=BackupData target=main>备份数据库</a>"
menu(6,3)="<a href=admin_data.asp?action=RestoreData target=main>恢复数据库</a>"
menu(6,4)="<a href=admin_data.asp?action=SpaceSize target=main>系统空间占用</a>"

'menu(6,1)="<a href=admin_CompressData.asp target=main>压缩数据库</a>"
'menu(6,2)="<a href=admin_BackupData.asp target=main>备份数据库</a>"
'menu(6,3)="<a href=admin_RestoreData.asp target=main>恢复数据库</a>"
'menu(6,4)="<a href=admin_SpaceSize.asp target=main>系统空间占用</a>"

menu(7,0)="文件管理"
menu(7,1)="<a href=admin_upUserface.asp target=main>上传头像管理</a>"
menu(7,2)="<a href=admin_uploadlist.asp target=main>上传文件管理</a>"
menu(7,3)="<a href=admin_bbsface.asp?orders=1 target=main>注册头像管理</a>"
menu(7,4)="<a href=admin_bbsface.asp?orders=2 target=main>发贴心情管理</a>"
menu(7,5)="<a href=admin_bbsface.asp?orders=3 target=main>发贴表情管理</a>"
%>
<BODY leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<table width=100% cellpadding=0 cellspacing=0 border=0 align=left>
    <tr><td valign=top>
<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
    <td height=42 valign=bottom>
	  <img src="skin/default/title.gif" width=158 height=38>
    </td>
  </tr>
</table>
<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background=skin/default/title_bg_quit.gif  >
	  <span><a href="admin_index.asp" target=_top><b>管理首页</b></a> | <a href=admin_logout.asp target=_top><b>退出</b></a></span>
    </td>
  </tr>
</table>
&nbsp;
<%
	dim j
	dim tmpmenu
	dim menuname
	dim menurl
for i=0 to ubound(menu,1)
%>
<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="skin/default/admin_left_<%=i+1%>.gif" id=menuTitle1 onclick="showsubmenu(<%=i%>)">
	  <span><%=menu(i,0)%></span>
	</td>
  </tr>
  <tr>
    <td style="display:" id='submenu<%=i%>'>
<div class=sec_menu style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=135>
	<%
	for j=1 to ubound(menu,2)
	if isempty(menu(i,j)) then exit for
	%>
<tr><td height=20><%=menu(i,j)%></td></tr>
<%
	next
%>
</table>
	  </div>
<div  style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=135>
<tr><td height=20></td></tr>
</table>
	  </div>
	</td>
  </tr>
</table>
<%next%>
&nbsp;
<table cellpadding=0 cellspacing=0 width=158 align=center>
  <tr>
    <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="skin/default/admin_left_9.gif" id=menuTitle1>
	  <span>动网论坛信息</span>
	</td>
  </tr>
  <tr>
    <td>
<div class=sec_menu style="width:158">
<table cellpadding=0 cellspacing=0 align=center width=135>

<tr><td height=20>&nbsp;<br><a href="http://www.aspsky.net/" target=_blank>版权所有：<BR>动网先锋(<font face=Verdana, Arial, Helvetica, sans-serif size=1>AspSky<font color=#CC0000>.Net</font></font>)</a><BR>
<a href="http://www.dvbbs.net/" target=_blank>支持论坛：<BR>动网论坛(<font face=Verdana, Arial, Helvetica, sans-serif size=1>Dvbbs<font color=#CC0000>.Net</font></font>)</a><BR><BR>
</td></tr>

</table>
	  </div>
	</td>
  </tr>
</table>
&nbsp;
<%
end sub

sub admin_login()
	stats="论坛管理登陆"
	if not founduser then
		Errmsg="<li>您不是系统管理员！"
		founderr=true
	end if
	if founderr then
		call nav()
		call head_var(2,0,"","")
		call dvbbs_error()
		call footer()
	else
		if request("reaction")="chklogin" then
			call chklogin()
			if founderr then
				call nav()
				call head_var(2,0,"","")
				call Dvbbs_error()
				call footer()
			end if
		else
			dim num1
			dim rndnum
			Randomize
			Do While Len(rndnum)<4
			num1=CStr(Chr((57-48)*rnd+48))
			rndnum=rndnum&num1
			loop
			session("verifycode")=rndnum
			call admin_login_main()
		end if
	end if
end sub

sub admin_login_main()
%>
<html>
<head>
<meta NAME=GENERATOR Content="Microsoft FrontPage 4.0" CHARSET=GB2312>
<meta name=keywords content='动网先锋,动网论坛,dvbbs'>
<title><%=Forum_info(0)%>--<%=stats%></title>
<link rel=stylesheet href="Forum_admin.css" type=text/css>
</head>
<body leftmargin=0 bottommargin=0 rightmargin=0 topmargin=0 marginheight=0 marginwidth=0>
<title><%=Forum_info(0)%>--<%=stats%></title>
<p>&nbsp;</p>
<p>&nbsp;</p>
<form action="admin_index.asp?action=admin_login&reaction=chklogin" method=post> 
<table style="width:65%" border=0 cellspacing=1 cellpadding=3  align=center class=tableBorder>
    <tr>
    <th bgcolor=<%=Forum_body(2)%> valign=middle colspan=2 height=25><%=Forum_info(0)%>管理登陆</th></tr>
    <tr>
    <td valign=middle class=Forumrow>请输入您的用户名</td>
    <td valign=middle class=Forumrow><INPUT name=username type=text></td></tr>
    <tr>
    <td valign=middle class=Forumrow>请输入您的密码</font></td>
    <td valign=middle class=Forumrow><INPUT name=password type=password></td></tr>
    <tr>
    <td valign=middle class=Forumrow>请输入您的附加码</td>
    <td valign=middle class=Forumrow><INPUT name=verifycode type=text> &nbsp; 请在附加码框输入  <b><span style="background-color: #FFFFFF"><font color=#000000><%=session("verifycode")%></font></span></b></td></tr>
    <tr>
    <td valign=middle colspan=2 align=center class=forumRow><input type=submit name=submit value="登 陆"></td></tr></table></form>
</body>
</html>
<%
call footer()
end sub

sub chklogin()
	username=trim(replace(request("username"),"'",""))
	password=md5(trim(replace(request("password"),"'","")))
	if request("verifycode")="" then
		errmsg=errmsg+"<br>"+"<li>请返回输入确认码。<li><b>返回后请刷新登陆页面后重新输入正确的信息。</b>"
		founderr=true
	elseif session("verifycode")="" then
		errmsg=errmsg+"<br>"+"<li>请不要重复提交，如需重新登陆请返回登陆页面。<li><b>返回后请刷新登陆页面后重新输入正确的信息。</b>"
		founderr=true
	elseif session("verifycode")<>trim(request("verifycode")) then
		errmsg=errmsg+"<br>"+"<li>您输入的确认码和系统产生的不一致，请重新输入。<li><b>返回后请刷新登陆页面后重新输入正确的信息。</b>"
		founderr=true
	end if
	session("verifycode")=""
	if username="" or password="" then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>请输入您的用户名或密码。<li><b>返回后请刷新登陆页面后重新输入正确的信息。</b>"
	end if
	if founderr then exit sub
	ip=Request.ServerVariables("REMOTE_ADDR")
	set rs=conn.execute("select * from admin where username='"&username&"' and adduser='"&membername&"'")
	if rs.eof and rs.bof then
		rs.close
		set rs=nothing
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>您输入的用户名和密码不正确或者您不是系统管理员。<br><li>请<a href=admin_login.asp>重新输入</a>您的密码。<li><b>返回后请刷新登陆页面后重新输入正确的信息。</b>"
		exit sub
	else
		if trim(rs("password"))<>password then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>您输入的用户名和密码不正确或者您不是系统管理员。<br><li>请<a href=admin_login.asp>重新输入</a>您的密码。<li><b>返回后请刷新登陆页面后重新输入正确的信息。</b>"
		exit sub
		else
		session("flag")=rs("flag")
		session.timeout=45
		conn.execute("update admin set LastLogin=Now(),LastLoginIP='"&ip&"' where username='"&username&"'")
		rs.close
		set rs=nothing
		response.write "<script>location.href='admin_index.asp'</script>"
		end if
	end if
end sub

sub admin_main()
%>
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<%
if not master or session("flag")="" then
	Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登陆</a>后进入。"
	call dvbbs_error()
else
	Dim theInstalledObjects(17)
    theInstalledObjects(0) = "MSWC.AdRotator"
    theInstalledObjects(1) = "MSWC.BrowserType"
    theInstalledObjects(2) = "MSWC.NextLink"
    theInstalledObjects(3) = "MSWC.Tools"
    theInstalledObjects(4) = "MSWC.Status"
    theInstalledObjects(5) = "MSWC.Counters"
    theInstalledObjects(6) = "IISSample.ContentRotator"
    theInstalledObjects(7) = "IISSample.PageCounter"
    theInstalledObjects(8) = "MSWC.PermissionChecker"
    theInstalledObjects(9) = "Scripting.FileSystemObject"
    theInstalledObjects(10) = "adodb.connection"
    
    theInstalledObjects(11) = "SoftArtisans.FileUp"
    theInstalledObjects(12) = "SoftArtisans.FileManager"
    theInstalledObjects(13) = "JMail.SMTPMail"
    theInstalledObjects(14) = "CDONTS.NewMail"
    theInstalledObjects(15) = "Persits.MailSender"
    theInstalledObjects(16) = "LyfUpload.UploadFile"
    theInstalledObjects(17) = "Persits.Upload.1"
%>
<table cellpadding=3 cellspacing=1 border=0 width=95% align=center>
<tr>
<td width="100%" valign=top>
<p><b>欢迎光临 动网论坛 管理控制面板</b>&nbsp;<%=Version%><br><BR>
官方网站：<a href="http://www.aspsky.net/" target="_blank">动网先锋</a><br>
官方论坛：<a href="http://www.dvbbs.net/" target="_blank">动网论坛</a>
</p>
在这里，您可以控制你所有的论坛设置。请在此页的左侧选择您要进行管理的链接。<br>
友情提示：用户权限的优先顺序为 用户组默认设置－－论坛用户组权限设置－－(分)论坛用户权限设置
</td>
</tr>
</table>
<table cellpadding="2" cellspacing="1" border="0" width="95%" class="tableBorder" align=center>
<tr><th class="tableHeaderText" colspan=2 height=25>论坛信息统计</th><tr>
<tr>
<td width="50%"  class="forumRow" height=23>服务器类型：<%=Request.ServerVariables("OS")%>(IP:<%=Request.ServerVariables("LOCAL_ADDR")%>)</td>
<td width="50%" class="forumRowHighlight">脚本解释引擎：<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
</tr>
<tr>
<td width="50%" class="forumRow" height=23>站点物理路径：<%=request.ServerVariables("APPL_PHYSICAL_PATH")%></td>
<td width="50%" class="forumRowHighlight">数据库地址：</td>
</tr>
<tr>
<td width="50%" class="forumRow" height=23>FSO文本读写：<%If Not IsObjInstalled(theInstalledObjects(9)) Then%><font color="<%=Forum_body(8)%>"><b>×</b></font><%else%><b>√</b><%end if%></td>
<td width="50%" class="forumRowHighlight">数据库使用：<%If Not IsObjInstalled(theInstalledObjects(10)) Then%><font color="<%=Forum_body(8)%>"><b>×</b></font><%else%><b>√</b><%end if%></td>
</tr>
<tr>
<td width="50%" class="forumRow" height=23>Jmail组件支持：<%If Not IsObjInstalled(theInstalledObjects(13)) Then%><font color="<%=Forum_body(8)%>"><b>×</b></font><%else%><b>√</b><%end if%></td>
<td width="50%" class="forumRowHighlight">CDONTS组件支持：<%If Not IsObjInstalled(theInstalledObjects(14)) Then%><font color="<%=Forum_body(8)%>"><b>×</b></font><%else%><b>√</b><%end if%></td>
</tr>
</table>

<p></p>

<table cellpadding="2" cellspacing="1" border="0" width="95%" class="tableBorder" align=center>
<tr><th class="tableHeaderText" colspan=2 height=25>论坛管理快捷方式</th><tr>
<FORM METHOD=POST ACTION="admin_user.asp?action=userSearch&userSearch=9"><tr>
<td width="20%"  class="forumRow" height=23>快速查找用户</td>
<td width="80%" class="forumRowHighlight">
<input type="text" name="username" size="30"><input type="submit" value="立刻查找">
<input type="hidden" name="userclass" value="0">
<input type="hidden" name="searchMax" value=100>
</td></FORM>
</tr>
<tr>
<td width="20%" class="forumRow" height=23>有用的链接</td>
<td width="80%" class="forumRowHighlight"><select onchange="jumpto(this.options[this.selectedIndex].value)">
		<option>>> 有用的链接 <<</option>
		<option value="http://www.dvbbs.net/">动网技术支持论坛</option>
		<option value="http://www.aspsky.net/article/">动网先锋文章区</option>
		<option value="http://www.dvbbs.net/vbscript/vbstoc.htm">Vbscript在线帮助文档</option>
		<option value="http://www.dvbbs.net/dvbbsinfo.asp">论坛变量查看文档</option>
</select></td>
</tr>
<tr>
<td width="20%" class="forumRow" height=23>快捷功能链接</td>
<td width="80%" class="forumRowHighlight"><a href=admin_board.asp?action=add>添加论坛类别</a> | <a href=admin_board.asp>管理论坛版面</a></td>
</tr>
<tr><form action="admin_update.asp?action=updat" method=post>
<td width="20%" class="forumRow" height=23>快速更新数据</td>
<td width="80%" class="forumRowHighlight">
<input type="submit" name="Submit" value="更新论坛数据">&nbsp;
<input type="submit" name="Submit" value="更新论坛总数据">
</td></form>
</tr>
</table>
<script language='javascript'> function jumpto(url) { if (url != '') { window.open(url); } } </script>
<p></p>

<table cellpadding="2" cellspacing="1" border="0" width="95%" class="tableBorder" align=center>
<tr><th class="tableHeaderText" colspan=2 height=25>动网论坛开发 & 贡献</th><tr>
<tr>
<td width="20%"  class="forumRow" height=23>产品负责</td>
<td width="80%" class="forumRowHighlight">
<a href="http://www.dvbbs.net/dispuser.asp?name=沙滩小子">沙滩小子</a>（联系：13088878365，QQ:421634，eway@aspsky.net）
</td>
</tr>
<tr>
<td width="20%" class="forumRow" height=23>商业发展</td>
<td width="80%" class="forumRowHighlight"><a href="http://www.dvbbs.net/dispuser.asp?name=quest">quest</a>（联系：13005069033，QQ:398196，quest@aspsky.net）</td>
</tr>
<tr>
<td width="20%" class="forumRow" height=23>开发人员</td>
<td width="80%" class="forumRowHighlight"><a href="http://www.dvbbs.net/dispuser.asp?name=沙滩小子">沙滩小子</a>，<a href="http://www.dvbbs.net/dispuser.asp?name=木鸟">木鸟</a>，<a href="http://www.dvbbs.net/dispuser.asp?name=fssunwin">fssunwin</a></td>
</tr>
<tr>
<td width="20%" class="forumRow" height=23>插件开发</td>
<td width="80%" class="forumRowHighlight">
动网论坛插件组织（Dvbbs Plus Organization）
</td>
</tr>
<tr>
<td width="20%" class="forumRow" height=23>美工设计</td>
<td width="80%" class="forumRowHighlight">
<font color=gray>暂无</font>
</td>
</tr>
</table>


<table align=center><BR>
Copyright (c) 2001-2002 <a href=http://www.aspsky.net><font face=Verdana, Arial, Helvetica, sans-serif size=1><b>AspSky<font color=#CC0000>.Net</font></b></font></a>. All Rights Reserved .<BR><BR><BR>
</td>
            </tr>
        </table>
<%
end if
end sub

Function IsObjInstalled(strClassString)
On Error Resume Next
IsObjInstalled = False
Err = 0
Dim xTestObj
Set xTestObj = Server.CreateObject(strClassString)
If 0 = Err Then IsObjInstalled = True
Set xTestObj = Nothing
Err = 0
End Function

sub admin_head()
%>
<title><%=Forum_info(0)%>--管理页面</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<BODY leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#D6DFF7">
<table width="100%" align=center cellpadding=0 cellspacing=0 border=0>
<tr>
<td height="23" class="forumHeaderBackgroundAlternate" valign=middle id=TableTitleLink>→ 欢迎<%=membername%>进入控制面板</td>
<td class="forumHeaderBackgroundAlternate" valign=middle id=TableTitleLink><a href="http://www.dvbbs.net/" target=_blank>动网会员交流</a></td>
<td class="forumHeaderBackgroundAlternate" align=right valign=middle id=TableTitleLink><a href="index.asp" target=_top>返回论坛首页</a></td>
</tr>
</table>
<%
end sub
%>