<!--#include file="conn.asp"-->
<!-- #include file="inc/char_board.asp" -->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/ubbcode.asp"-->
<%
dim AnnounceID
dim replyID
dim username
dim dateandtime
dim bgcolor,abgcolor
dim PostUserGroup
dim con,content
dim TotalUseTable
dim postbuyuser

if BoardID="" or not isInteger(BoardID) or BoardID=0 then
	Errmsg=Errmsg+"<br>"+"<li>错误的版面参数！请确认您是从有效的连接进入。"
	founderr=true
else
	BoardID=clng(BoardID)
end if
if request("id")="" then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>请指定相关贴子。"
elseif not isInteger(request("id")) then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>非法的贴子参数。"
else
	AnnounceID=request("id")
end if
if request("replyid")="" then
	replyid=Announceid
elseif not isInteger(request("replyid")) then
	replyid=Announceid
else
	replyid=request("replyid")
end if
if founderr then
	call nav()
	call head_var(2,0,"","")
	call dvbbs_error()
else
	stats="回复帖子"
	call nav()
	call head_var(1,BoardDepth,0,0)
	call main()
	if founderr then call dvbbs_error()
end if
call footer()

sub main()
if Cint(Board_Setting(43))=1 then
	Errmsg=Errmsg+"<br>"+"<li>本论坛已经被管理员限制了不允许发贴。"
	founderr=true
	exit sub
end if
if cint(Board_Setting(1))=1 then
	if Cint(GroupSetting(37))=0 then
		Errmsg=ErrMsg+"<Br>"+"<li>您没有权限进入隐含论坛！"
		founderr=true
		exit sub
	end if
end if
if cint(Board_Setting(2))=1 then
	if not founduser then
		Errmsg=Errmsg+"<br>"+"<li>本论坛为认证论坛，请<a href=login.asp>登陆</a>并确认您的用户名已经得到管理员的认证后进入。"
		founderr=true
		exit sub
	else
		if chkboardlogin(boardid,membername)=false then
		Errmsg=Errmsg+"<br>"+"<li>本论坛为认证论坛，请确认您的用户名已经得到管理员的认证后进入。"
		founderr=true
		exit sub
		end if
	end if
end if

set rs=conn.execute("select PostTable,locktopic from topic where TopicID="&AnnounceID)
if not (rs.eof and rs.bof) then
	if rs(1)=1 or rs(1)=2 then
		Errmsg=Errmsg+"<br>"+"<li>本主题已被锁定或者删除，不能发表回复。"
		founderr=true
		exit sub
	end if
	TotalUseTable=rs(0)
else
	ErrMsg=ErrMsg+"<br>"+"<li>您指定的贴子不存在</li>"
	founderr=true
	exit sub
end if

if replyid=Announceid then
	set rs=conn.execute("select top 1 Announceid from "&TotalUseTable&" where rootID="&AnnounceID&" order by Announceid")
	if not(rs.bof and rs.eof) then
		replyID=rs(0)
	else
		ErrMsg=ErrMsg+"<br>"+"<li>您指定的贴子不存在</li>"
		founderr=true
		exit sub
	end if
end if

sql="select b.body,b.topic,b.locktopic,b.username,b.dateandtime,b.isbest,u.lockuser,u.UserGroupID from "&TotalUseTable&" b inner join [user] u on b.postuserid=u.userid where b.AnnounceID="&replyID
set rs=conn.execute(sql)
if rs.eof and rs.bof then
	Errmsg=Errmsg+"<br>"+"<li>没有找到相关帖子。"
	founderr=true
	exit sub
else
	if rs("locktopic")=1 or rs("locktopic")=2 then
		Errmsg=Errmsg+"<br>"+"<li>本主题已被锁定或者删除，不能发表回复。"
		founderr=true
		exit sub
	elseif rs("lockuser")=2 then
   		con=""
	elseif rs("isbest")=1 and Cint(GroupSetting(41))=0 then
		con=""
	else
   		con=rs("body")
	end if
	username=rs("username")
	dateandtime=rs("dateandtime")
	PostUserGroup=rs("UserGroupID")
	if username=membername then
		if Cint(GroupSetting(4))=0 then
		Errmsg=Errmsg+"<br>"+"<li>您没有在本论坛回复自己主题的权限。"
		founderr=true
		exit sub
		end if
	else
		if Cint(GroupSetting(5))=0 then
		Errmsg=Errmsg+"<br>"+"<li>您没有在本论坛回复其他人帖子的权限，请<a href=login.asp>登陆</a>或者同管理员联系。"
		founderr=true
		exit sub
		end if
		if Cint(GroupSetting(2))=0 then
		Errmsg=Errmsg+"<br>"+"<li>您没有浏览在本论坛查看其他人发布的帖子的权限，请<a href=login.asp>登陆</a>或者同管理员联系。"
		founderr=true
		exit sub
		end if
	end if
end if
set rs=nothing

if request("reply")="true" then
	con = reubbcode(con)
	con = replace(con, "&gt;", ">")
	con = replace(con, "&lt;", "<")
	con = Replace(con, "</P><P>", CHR(10) & CHR(10))
	con = Replace(con, "<BR>", CHR(10))
	content = "[quote][b]以下是引用[i]"&username&"在"&dateandtime&"[/i]的发言：[/b]" & chr(13)
	content = content & con & chr(13)
	content = content & "[/quote]"
else
	content=""
end if
%>
<script src="inc/ubbcode.js"></script>
<form action="SaveReAnnounce.asp?method=Topic&boardID=<%=request("boardid")%>" method=POST onSubmit="submitonce(this)" name=frmAnnounce>
<input type="hidden" name="upfilerename">
<input type=hidden name="followup" value="<%=replyID%>">
<input type=hidden name="rootID" value="<%=AnnounceID%>">
<input type=hidden name="star" value="<%=request("star")%>">
<input type=hidden name="TotalUseTable" value="<%=TotalUseTable%>">

<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
    <tr>
      <th width=100% align=left colspan=2>&nbsp;&nbsp;发表回复<%if isaudit=1 then%>（本版面所有帖子都要经过管理员审核方可发表）<%end if%></th>
    </tr>
        <tr>
          <td width=20% class=tablebody2><b>用户名</b></td>
          <td width=80% class=tablebody2><input name=username value=<%=membername%> class=FormClass>&nbsp;&nbsp;<font color=<%=Forum_body(8)%>><b>*</b></font><a href=reg.asp>您没有注册？</a> 
          </td>
        </tr>
        <tr>
          <td width=20% class=tablebody1><b>密码</b></td>
          <td width=80% class=tablebody1><input name=passwd type=password value=<%=htmlencode(memberword)%> class=FormClass><font color=<%=Forum_body(8)%>>&nbsp;&nbsp;<b>*</b></font><a href=lostpass.asp>忘记密码？</a></td>
        </tr>
        <tr>
          <td width=20% class=tablebody2><b>主题标题</b>
              <SELECT name=font onchange=DoTitle(this.options[this.selectedIndex].value)>
              <OPTION selected value="">选择话题</OPTION> <OPTION value=[原创]>[原创]</OPTION> 
              <OPTION value=[转帖]>[转帖]</OPTION> <OPTION value=[灌水]>[灌水]</OPTION> 
              <OPTION value=[讨论]>[讨论]</OPTION> <OPTION value=[求助]>[求助]</OPTION> 
              <OPTION value=[推荐]>[推荐]</OPTION> <OPTION value=[公告]>[公告]</OPTION> 
              <OPTION value=[注意]>[注意]</OPTION> <OPTION value=[贴图]>[贴图]</OPTION>
              <OPTION value=[建议]>[建议]</OPTION> <OPTION value=[下载]>[下载]</OPTION>
              <OPTION value=[分享]>[分享]</OPTION></SELECT>
	  </td>
	<td width=80% class=tablebody2><input name=subject size=70 maxlength=80 class=FormClass>&nbsp;&nbsp;<font color=<%=Forum_body(8)%>><b>*</b></font>不得超过 25 个汉字或50个英文字符
	 </td>
        </tr>
        <tr>
          <td width=20% valign=top class=tablebody1><b>当前心情</b><br><li>将放在帖子的前面</td>
          <td width=80% class=tablebody1>
<%for i=0 to Forum_PostFaceNum-1%>
	<input type="radio" value="<%=Forum_PostFace(i)%>" name="Expression" <%if i=0 then response.write "checked"%>><img src="<%=forum_info(8)&Forum_PostFace(i)%>" >&nbsp;&nbsp;
<%if i>0 and ((i+1) mod 9=0) then response.write "<br>"%>
<%next%>
 </td>
        </tr>
<%if Cint(GroupSetting(7))=1 or Cint(GroupSetting(7))=3 then%>
<tr>
<td width=20% valign=top class=tablebody2>
<b>文件上传</b>
<select name="filetype" size=1>
<option value="">允许类型</option>
<%
dim iupload
iupload=split(Board_Setting(19),"|")
for i=0 to ubound(iupload)
response.write "<option value="&iupload(i)&">"&iupload(i)&"</option>"
next
%>
</select>
</td><td width=80% class=tablebody2>
<iframe name="ad" frameborder=0 width=100% height=24 scrolling=no src=saveannounce_upload.asp?boardid=<%=boardid%>></iframe>
</td></tr>
<%end if%>
        <tr>
          <td width=20% valign=top class=tablebody1>
<b>内容</b><br>
<li>HTML标签： <%=iif(Board_Setting(5),"可用","不可用")%>
<li>UBB标签： <%=iif(Board_Setting(6),"可用","不可用")%>
<li>贴图标签： <%=iif(Board_Setting(7),"可用","不可用")%>
<li>多媒体标签：<%=iif(Board_Setting(9),"可用","不可用")%>
<li>表情字符转换：<%=iif(Board_Setting(8),"可用","不可用")%>
<li>上传图片：<%=iif(Forum_Setting(3),"可用","不可用")%>
<li>最多<%=Board_Setting(16)\1024%>KB<BR><BR>
<B>特殊内容</B><BR>
<li><%=iif(Board_Setting(10),"<a href=""javascript:money()"" title=""使用语法：[money=可浏览该部分内容需要金钱数]内容[/money]"">金钱帖</a>","金钱帖：不可用")%>
<li><%=iif(Board_Setting(11),"<a href=""javascript:point()"" title=""使用语法：[point=可浏览该部分内容需要积分数]内容[/point]"">积分帖</a>","积分帖：不可用")%>
<li><%=iif(Board_Setting(12),"<a href=""javascript:usercp()"" title=""使用语法：[usercp=可浏览该部分内容需要魅力数]内容[/usercp]"">魅力帖</a>","魅力帖：不可用")%>
<li><%=iif(Board_Setting(13),"<a href=""javascript:power()"" title=""使用语法：[power=可浏览该部分内容需要威望数]内容[/power]"">威望帖</a>","威望帖：不可用")%>
<li><%=iif(Board_Setting(14),"<a href=""javascript:article()"" title=""使用语法：[post=可浏览该部分内容需要文章数]内容[/post]"">文章帖</a>","文章帖：不可用")%>
<li><%=iif(Board_Setting(15),"<a href=""javascript:replyview()"" title=""使用语法：[replyview]该部分内容回复后可见[/replyview]"">回复帖</a>","回复帖：不可用")%>
<li><%=iif(Board_Setting(23),"<a href=""javascript:usemoney()"" title=""使用语法：[usemoney=浏览该部分内容需要消耗的金钱数]内容[/usemoney]"">出售帖</a>","出售帖：不可用")%>
	  </td>
          <td width=80% class=tablebody1>
<%if Cint(Board_Setting(4)) then%>
<!--#include file="inc/getubb.asp"-->
<%end if%>
<textarea class=smallarea cols=95 name=Content rows=12 wrap=VIRTUAL title=可以使用Ctrl+Enter直接提交贴子 class=FormClass onkeydown=ctlent()><%=server.htmlencode(content)%></textarea>
          </td>
        </tr>

		<tr>
                <td class=tablebody1 valign=top colspan=2 style="table-layout:fixed; word-break:break-all"><b>点击表情图即可在帖子中加入相应的表情</B><br>&nbsp;
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
                <td valign=top class=tablebody1><b>选项</b></td>
                <td valign=middle class=tablebody1><input type=checkbox name=signflag value=yes checked>是否显示您的签名？<br>
                <input type=checkbox name=emailflag value=yes>有回复时使用邮件通知您？</font>
<BR><BR></td>
                </tr><tr>
                <td valign=middle colspan=2 align=center class=tablebody2>
                <input type=Submit value='发 表' name=Submit> &nbsp; <input type=button value='预 览' name=Button onclick=gopreview()>&nbsp;
<input type=reset name=Submit2 value='清 除'>
                </td></form></tr>
      </table>
</form>
<form name=preview action=preview.asp?boardid=<%=boardid%> method=post target=preview_page>
<input type=hidden name=title value=><input type=hidden name=body value=>
</form>
<hr width=<%=Forum_body(12)%> size=1>
<%
	sql="Select top 10 b.UserName,b.Topic,b.dateandtime,b.body,b.announceid,b.isbest,u.lockuser,u.UserGroupID,b.postbuyuser from "&TotalUseTable&" b inner join [user] u on b.postuserid=u.userid where b.boardid="&boardid&" and b.rootid="&Announceid&" and b.locktopic<2 order by b.announceid desc"
	set rs=conn.execute(sql)
	do while not rs.eof
	PostUserGroup=rs("UserGroupID")
	postbuyuser=rs("postbuyuser")
	if bgcolor="tablebody1" then
		bgcolor="tablebody2"
		abgcolor="tablebody1"
	else
		bgcolor="tablebody1"
		abgcolor="tablebody2"
	end if
%>
<TABLE border=0 width=<%=Forum_body(12)%> align=center>
  <TBODY>
  <TR>
    <TD valign=middle align=top class=<%=bgcolor%>>
--&nbsp;&nbsp;作者：<%=htmlencode(rs("username"))%><br>
--&nbsp;&nbsp;发布时间：<%=rs("dateandtime")%><br><br>
--&nbsp;&nbsp;<%=htmlencode(rs("topic"))%><br>
<%
	if rs("lockuser")=2 then
		response.write "该用户发言已被屏蔽"
	elseif rs("isbest")=1 and Cint(GroupSetting(41))=0 then
		response.write "您没有浏览该精华帖子的权限"
	else
		response.write dvbcode(rs("body"),PostUserGroup,1)
	end if
%>
    </TD></TR></TBODY></TABLE> 
	<hr size=1 class=tableborder1>
<%
	rs.movenext
	loop	
	set rs=nothing
	call activeonline()
end sub
%>