<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/char_board.asp" -->
<%
dim AnnounceID
dim replyID
dim username
dim rs_old
dim old_user
dim con,content
dim topic
dim olduser,oldpass
dim totalusetable
dim CanEditPost
CanEditPost=False
if BoardID="" or not isInteger(BoardID) or BoardID=0 then
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
if request("id")="" then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>请指定相关贴子。"
elseif not isInteger(request("id")) then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>非法的贴子参数。"
else
	AnnounceID=request("id")
end if
if request("replyID")="" then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>请指定相关贴子。"
elseif not isInteger(request("replyID")) then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>非法的贴子参数。"
else
	replyID=request("replyID")
end if
if not founduser then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>请登陆后进行操作。"
end if
if Cint(Board_Setting(43))=1 then
	Errmsg=Errmsg+"<br>"+"<li>本论坛已经被管理员限制了不允许发贴。"
	founderr=true
end if
if cint(Board_Setting(1))=1 then
	if Cint(GroupSetting(37))=0 then
		Errmsg=ErrMsg+"<Br>"+"<li>您没有权限进入隐含论坛！"
		founderr=true
	end if
end if
stats=boardtype & "编辑帖子"
if founderr then
	call nav()
	call head_var(2,0,"","")
	call dvbbs_error()
else
	call nav()
	call head_var(1,BoardDepth,0,0)
	call main()
	if founderr then call dvbbs_error()
end if
call activeonline()
call footer()

sub main()
set rs=conn.execute("select PostTable from topic where TopicID="&AnnounceID)
if rs.eof and rs.bof then
	Errmsg=Errmsg+"<br>"+"<li>没有找到相应的帖子。"
	Founderr=true
	exit sub
else
	TotalUseTable=rs(0)
end if

sql="select b.username,b.topic,b.body,b.dateandtime,u.usergroupID from "&TotalUseTable&" b,[user] u where b.postuserid=u.userid and b.AnnounceID="&replyID
set rs=conn.execute(sql)
if rs.eof and rs.bof then
	Errmsg=Errmsg+"<br>"+"<li>没有找到相应的帖子。"
	Founderr=true
	exit sub
else
	topic=rs("topic")
	con=rs("body")
	old_user=rs("username")
	if Clng(forum_setting(51))>0 and not (master or boardmaster or superboardmaster) then
		if Datediff("s",rs("dateandtime"),Now())>Clng(forum_setting(51))*60 then
		Errmsg=Errmsg+"<br>"+"<li>系统编辑帖子时限为"&forum_setting(51)&"分钟，而从您发表帖子到现在已经有"&Datediff("s",rs("dateandtime"),Now())\60&"分钟了。"
			founderr=true
			exit sub
		end if
	end if
	if rs("username")=membername then
		if Cint(GroupSetting(10))=0 then
			Errmsg=Errmsg+"<br>"+"<li>您没有在本论坛编辑自己帖子的权限，请<a href=login.asp>登陆</a>或者同管理员联系。"
			founderr=true
			exit sub
			CanEditPost=False
		else
			CanEditPost=True
		end if
	else
		if (master or superboardmaster or boardmaster) and Cint(GroupSetting(23))=1 then
			CanEditPost=True
		else
			CanEditPost=False
		end if
		if UserGroupID>3 and Cint(GroupSetting(23))=1 then
			CanEditPost=true
		end if
		if Cint(GroupSetting(23))=1 and FoundUserPer then
			CanEditPost=True
		elseif Cint(GroupSetting(23))=0 and FoundUserPer then
			CanEditPost=False
		end if
		if UserGroupID<4 and UserGroupID=rs("UserGroupID") then
			Errmsg=Errmsg+"<br>"+"<li>同等级用户不能修改。"
			Founderr=true
			exit sub
		elseif UserGroupID<4 and UserGroupID>rs("UserGroupID") then
			Errmsg=Errmsg+"<br>"+"<li>不能修改等级比您高的用户的帖子。"
			Founderr=true
			exit sub
		end if
		if not CanEditPost then
			Errmsg=Errmsg+"<br>"+"<li>您没有足够的权限编辑本帖子，请和管理员联系。"
			Founderr=true
			exit sub
		end if
	end if
end if

set rs=nothing

if con<>"" then
	content=replace(con,"<br>",chr(13))
	content=replace(content,"&nbsp;","")
	content=content+chr(13)
end if
%>
<script src="inc/ubbcode.js"></script>
<form action="SaveditAnnounce.asp?boardID=<%=boardID%>&replyID=<%=replyID%>&ID=<%=announceID%>" method=POST name=frmAnnounce>
  <input type=hidden name=followup value=<%=AnnounceID%>>
  <input type=hidden name="star" value="<%=request("star")%>">
  <input type=hidden name="TotalUseTable" value="<%=TotalUseTable%>">
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
    <tr>
      <th width=100% colspan=2>编辑帖子</th>
    </tr>
        <tr>
          <td width=20% class=tablebody1><b>用户名</b></td>
          <td width=80% class=tablebody1><input name=username value=<%=htmlencode(old_user)%> >&nbsp;&nbsp;<font color=<%=Forum_body(8)%>><b>*</b></font><a href=reg.asp>您没有注册？</a> 
          </td>
        </tr>
        <tr>
          <td width=20% class=tablebody1><b>密码</b></td>
          <td width=80% class=tablebody1><input name=passwd type=password value=<%=htmlencode(memberword)%> >&nbsp;&nbsp;<font color=<%=Forum_body(8)%>><b>*</b></font><a href=lostpass.asp>忘记密码？</a>版主编辑不需要密码</td>
        </tr>
        <tr>
          <td width=20% class=tablebody1><b>主题标题</b>
              <SELECT name=font onchange=DoTitle(this.options[this.selectedIndex].value)>
              <OPTION selected value="">选择话题</OPTION> <OPTION value=[原创]>[原创]</OPTION> 
              <OPTION value=[转帖]>[转帖]</OPTION> <OPTION value=[灌水]>[灌水]</OPTION> 
              <OPTION value=[讨论]>[讨论]</OPTION> <OPTION value=[求助]>[求助]</OPTION> 
              <OPTION value=[推荐]>[推荐]</OPTION> <OPTION value=[公告]>[公告]</OPTION> 
              <OPTION value=[注意]>[注意]</OPTION> <OPTION value=[贴图]>[贴图]</OPTION>
              <OPTION value=[建议]>[建议]</OPTION> <OPTION value=[下载]>[下载]</OPTION>
              <OPTION value=[分享]>[分享]</OPTION></SELECT>
         <td width=80% class=tablebody1><input name=subject size=70 maxlength=80 value=<%=htmlencode(Topic)%> >&nbsp;&nbsp;<font color=<%=Forum_body(8)%>><b>*</b></font>不得超过 25 个汉字或50个英文字符
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
<%if Cint(Board_Setting(4))=1 then%>
<!--#include file="inc/getubb.asp"-->
<%end if%>
<textarea class=smallarea cols=95 name=Content rows=12 wrap=VIRTUAL  onkeydown=ctlent()>
<%=server.htmlencode(content)%>
</textarea>
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
                <input type=checkbox name=emailbody value=yes>有回复时使用邮件通知您？
<BR><BR></td>
                </tr><tr>
                <td class=tablebody2 valign=middle colspan=2 align=center>
                <input type=Submit value="发 表" name=Submit> &nbsp; <input type=button value='预 览' name=Button onclick=gopreview()>&nbsp;<input type=reset name=Submit2 value="清 除">
                </td></form></tr>
      </table>
</form>
<form name=preview action=preview.asp?boardid=<%=boardid%> method=post target=preview_page>
<input type=hidden name=title value=><input type=hidden name=body value=>
</form>

<%
end sub
%>