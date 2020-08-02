<!--#include FILE="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/chkinput.asp" -->
<%
stats="个人上传管理"
call nav()

Dim TopicCount
Dim Pcount,endpage,star,page_count
if request("star")="" or not isnumeric(request("star")) then
	star=1
else
	star=clng(request("star"))
end if

if not founduser then
  	errmsg=errmsg+"<br>"+"<li>您没有<a href=login.asp target=_blank>登录</a>。"
	founderr=true
end if
if founderr=true then
	call head_var(2,0,"","")
	call dvbbs_error()
else
	call head_var(0,0,membername & "的控制面板","usermanager.asp")
	select case request("action")
	case "edit"
		call edit()
	case "fsave"
		call filesave()
	case "fadd"
		call addnew()	
	case "fsnew"
		call savenew()
	case "fdel"
		call fdel()
	case "alldel"
		call alldel()
	case else
		call main()
	end select
	if founderr then call dvbbs_error()
end if
call activeonline()
call footer()

sub main()
dim sname
dim stype
dim searchsql
stype=request("Stype")

TopicCount=filenum("")
if stype="" or not  isnumeric(stype)  then
	sname="所有文件"
	searchsql=""
	TopicCount=filenum("")
else
	select case stype
	case 1
	sname="图片集"
	searchsql="F_Type=1 and"
	case 2
	sname="FLASH集"
	searchsql="F_Type=2 and"
	case 3
	sname="音乐集"
	searchsql="F_Type=3 and"
	case 4
	sname="电影集"
	searchsql="F_Type=4 and"
	case 0
	sname="文件集"
	searchsql="F_Type=0 and"
	case else
	sname="所有文件"
	searchsql=""
	end select
	TopicCount=filenum(clng(stype))
end if
%>
<br>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr>
<th width=14% height=25 id=tabletitlelink><a href=usermanager.asp>控制面板首页</a></th>
<th width=14%  id=tabletitlelink><a href=mymodify.asp>基本资料修改</a></th>
<th width=14%  id=tabletitlelink><a href=modifypsw.asp>用户密码修改</a></th>
<th width=14%  id=tabletitlelink><a href=modifyadd.asp>联系资料修改</a></th>
<th width=14%  id=tabletitlelink><a href=usersms.asp>用户短信服务</a></th>
<th width=14%  id=tabletitlelink><a href=friendlist.asp>编辑好友列表</a></th>
<th width=14%  id=tabletitlelink><a href=favlist.asp>用户收藏管理</a></th>
</tr>
</table>
<br>
	<table cellpadding=3 cellspacing=1  align=center  class=tableborder2 >
	<form action="<%=ScriptName%>" method=get>
	<tr><td><a href=?action=fadd >［新增文件库］</a>，<a href=show.asp?username=<%=trim(membername)%> >［个人展区］</a>，目前共有<%=filenum("")%>个文件
<% if 	stype<>"" and isnumeric(stype) then %>
	，其中《<%=sname%>》共<font color=<%=Forum_body(8)%>><%=filenum(clng(stype))%></font>个。
<%end if
%>
	</td>
	<td width=200 align=right>
	<select name=Stype onchange='javascript:submit()'>
	<option value=all>查看文件类型
	<option value=0>文件集
	<option value=1>图片集
	<option value=2>FLASH集
	<option value=3>音乐集
	<option value=4>电影集
	<option value=5>所有文件
	</select>
	</td></tr></form>
	</table>
	<script language="JavaScript">
<!--
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
//-->
</script>
<%
dim F_Type,F_typename
'---------------------
response.write"<table cellpadding=3 cellspacing=1 align=center class=tableborder1>"&_
              "<form action="""&ScriptName&"?action=fdel""  method=post name=myfile><tr>"&_
              "<th colspan=7 height=25 align=left>-=> 最新上传文件</th></tr>"&_
            "<tr>"&_
                "<td align=center valign=middle width=30 class=tabletitle2><b>属性</b></td>"&_
                "<td align=center valign=middle width=* class=tabletitle2><b>文件说明</b></td>"&_
                "<td align=center valign=middle width=60 class=tabletitle2><b>大小</b></td>"&_
				"<td align=center valign=middle width=60 class=tabletitle2><b>下载/浏览</b></td>"&_
                "<td align=center valign=middle width=120 class=tabletitle2><b>日期</b></td>"&_
                "<td align=center valign=middle width=60 class=tabletitle2><b>类型</b></td>"&_
				"<td align=center valign=middle width=120 class=tabletitle2><b>操作</b></td>"&_
            "</tr>"
'---------------------
set rs=conn.execute("select * from [DV_Upfile] where  "&searchsql&"  F_UserID="&userid&" order by F_ID desc ")
if rs.eof and rs.bof then
	response.write"<tr><td class=tablebody1 align=center valign=middle colspan=7>您的文件库中没有任何内容。</td></tr>"
else
		 	if TopicCount mod Cint(Forum_Setting(11))=0 then
     		Pcount= TopicCount \ Cint(Forum_Setting(11))
  			else
     		Pcount= TopicCount \ Cint(Forum_Setting(11))+1
  			end if
			RS.MoveFirst
			if star > Pcount then star = Pcount
			if star < 1 then star = 1
			RS.Move (star-1) * Forum_Setting(11)
			page_count=0

		do while not rs.eof and page_count < Cint(Forum_Setting(11))
F_Type=rs("F_Type")

if F_Type=1 then
	F_typename="图片集"
elseif F_Type=2 then
	F_typename="FLASH集"
elseif F_Type=3 then
	F_typename="音乐集"
elseif F_Type=4 then
	F_typename="电影集"
else
	F_typename="文件集"
end if
response.write	"<tr>"
response.write	"<td align=center valign=middle  class=tablebody2><img src=""images/files/"&rs("F_FileType")&".gif"" border=0></td>"
response.write	"<td align=left valign=middle  class=tablebody1><a href=""fileshow.asp?boardid="&rs("F_BoardID")&"&id="&rs("F_ID") &"""  target=""_blank"" >"
if rs("F_Readme")<>"" or not isnull(rs("F_Readme")) then
if len(rs("F_Readme"))>26 then
response.write ""&left(htmlencode(replace(rs("F_Readme"),chr(10)," ")),26)&"..."
else
response.write Server.htmlencode(rs("F_Readme"))
end if
end if
response.write"</a></td>"
response.write	"<td align=right valign=middle  class=tablebody2>"&rs("F_FileSize")&" <b>B</b></td>"
response.write	"<td align=center valign=middle  class=tablebody1>"&rs("F_DownNum")&"/"&rs("F_ViewNum")&"</td>"
response.write	"<td align=left valign=middle  class=tablebody2>"&rs("F_AddTime")&"</td>"
response.write	"<td align=center valign=middle  class=tablebody1>"&F_typename&"</td>"
response.write	"<td align=center valign=middle  class=tablebody2>"

if GroupSetting(48)=1 then
	if instr(rs("F_Filename"),"/") and cint(rs("F_Flag"))=1 then
response.write	" <input type=checkbox name=delid value="""&rs("F_ID")&"""> "
else
response.write	" <input type=checkbox  value="""" disabled > "
	end if
response.write"<a href=?action=edit&editid="&rs("F_ID")&"  >编辑</a>  | <a href=fileshow.asp?action=send&id="&rs("F_ID")&"  >发送</a>"
else
response.write	" —— "
end if
response.write	"</td>"
response.write	"</tr>"
page_count = page_count + 1
		rs.movenext
		loop
end if
rs.close
set rs=nothing
response.write "<tr><td  colspan=7 align=right class=tablebody2>请选择可以删除的文件，<input type=checkbox name=chkall value=on onclick=""CheckAll(this.form)"">全选 <input type=submit name=Submit value=执行  onclick=""{if(confirm('您确定执行的操作吗?')){this.document.myfile.submit();return true;}return false;}""></td></tr>"
response.write"</form></table>"

call list()
end sub

'分页代码
sub list()
response.write "<table cellpadding=0 cellspacing=3 border=0 width="&Forum_body(12)&" align=center><form method=post action="&ScriptName&" ><tr><td valign=middle nowrap align=left>"&stats&"文件共<b>"&TopicCount&"</b>个，共<b><font color=red>"&Pcount&"</font></b>页</td><td valign=middle  align=right>"

if star > 4 then
	response.write "<a href=""?&star=1"">[1]</a> ..."
end if
if Pcount>star+3 then
	endpage=star+3
else
	endpage=Pcount
end if
for i=star-3 to endpage
	if not i<1 then
		if i = clng(star) then
        response.write " <font color=#996633>["&i&"]</font>"
		else
        response.write " <a href=""?star="&i&""">["&i&"]</a>"
		end if
	end if
next
if star+3 < Pcount then 
	response.write "... <a href=""?star="&Pcount&""">["&Pcount&"]</a>"
end if
response.write " 转到:<input type=text name=star size=3 maxlength=10  value='"& star &"'><input type=submit value=Go  id=button1 name=button1 >"     
response.write "</td></tr></form></table>"
end sub

'编辑文件
sub edit()
dim editid
dim F_Type,F_typename,filename,chefile,con,body
dim F_Username,postid,F_rootid,F_bbsid
dim myurl

myurl=false
editid=trim(request("editid"))
if not isnumeric(editid) or isnull(editid) then 
	errmsg=errmsg+"<br>"+"<li>执行的数据不存在。"
	founderr=true
	exit sub
end if
set rs=conn.execute("select * from [DV_Upfile] where F_ID="&editid)
if rs.eof and rs.bof then
	errmsg=errmsg+"<br>"+"<li>编辑的文件已不存在。"
	founderr=true
	exit sub
else
	F_Username=rs("F_Username")
	F_Type=rs("F_Type")
	filename=rs("F_Filename")
	con=rs("F_Readme")

	if instr(filename,"/")=0 then
	myurl=true
	filename="UploadFile/"&filename
	end if
	if F_Type=1  then
	F_typename="<img src='"&filename&"'  style='border: 1 solid #000000' width=80 height=80 >"
	else
	F_typename=rs("F_FileType")&"文件"
	end if

	if con<>"" then
	body=replace(con,"<br>",chr(13))
	body=replace(body,"&nbsp;","")
	body=body+chr(13)
	end if
	
%>
 <table cellpadding=3 cellspacing=1 align=center class=tableborder1>
 <form action="?action=fsave" method=post >
	<tr>
	<th width="100%" height="26" colspan="2" align=left>编辑<%=F_Username%>上传文件</th>
	</tr>
	<tr>
	<td width="15%" height="26" class=tablebody2 align=right>编辑的文件：</td>
	<td width="85%"  height="26" class=tablebody1 align=left>
	<a href="fileshow.asp?boardid=<%=rs("F_BoardID")%>&id=<%=rs("F_ID") %>"  target="_blank" ><%=F_typename%></a>
	</td>
	</tr>
<% if myurl=false then %>
	<tr>
	<td width="15%" height="26" class=tablebody2 align=right>文件地址：</td>
	<td width="85%"  height="26" class=tablebody1 align=left><input type=text name="fileurl" size="80%"  value="<%=rs("F_Filename")%>">
	</td>
	</tr>
<% else 
	if instr(rs("F_AnnounceID"),"|") then
		postid=split(rs("F_AnnounceID"),"|")
		F_rootid=postid(0)
		F_bbsid=postid(1)
	end if
%>
	<tr>
	<td valign=middle align=right  class=tablebody2 >文件帖子系数：</td>
	<td width="100%" height="26" class=tablebody1>
	<% if (master or superboardmaster or boardmaster) and GroupSetting(48)=1 then%>
	论坛ID:<input type=text name="F_BoardID" value="<%= rs("F_BoardID") %>" size=8>&nbsp;
	帖子ID:<input type=text name="F_AnnounceID" value="<%= rs("F_AnnounceID") %>">&nbsp;<font color=<%=Forum_body(8)%>>（若文件与帖子内容未有改动,请不要修改相关系数。）</font>
	<%else%>
	<input type=hidden name="F_BoardID" value="<%= rs("F_BoardID") %>" >
	<input type=hidden name="F_AnnounceID" value="<%= rs("F_AnnounceID") %>">
	<%end if%>
	<a href="dispbbs.asp?Boardid=<%=rs("F_BoardID")%>&ID=<%=F_Rootid%>&replyID=<%=F_bbsid%>&skin=1" target="_blank" title="浏览相关帖子">[<font color=<%=Forum_body(8)%>>查看相关讨论</font>]</a>
	</td>
    	</tr>
<%end if%>
	<tr>
	<td width="15%" height="26" class=tablebody2 align=right>文件的说明：</td>
	<td width="85%"  height="26" class=tablebody1 align=left>
	<textarea name="F_Readme" rows=6 cols=80%  wrap=VIRTUAL><%response.write Server.htmlencode(body)%></textarea>
	<input type=hidden name="saveid" value="<%= rs("F_ID") %>">　（不得超过１２５个汉字或２５０个字符）
	</td>
	</tr>
	 <%if myurl=true then %>
	<tr>
	<td valign=middle align=right  class=tablebody2 >文件允许删除：</td>
	<td valign=middle class=tablebody1>
	  <input type=radio name=Fflag value=0 <%if rs("F_Flag")=0 then %> checked<%end if%>>是&nbsp;
	  <input type=radio name=Fflag value=2 <%if rs("F_Flag")=2 then %> checked<%end if%>>否
	</td>
	</tr>
	<%else%>
	<tr>
	<td valign=middle align=right  class=tablebody2 >文件允许删除：</td>
	<td valign=middle class=tablebody1>
	  <input type=radio name=Fflag value=1 <%if rs("F_Flag")=1 then %> checked<%end if%>>是&nbsp;
	  <input type=radio name=Fflag value=2 <%if rs("F_Flag")=2 then %> checked<%end if%>>否
	</td>
	</tr>
	<%end if%>
	<tr>
	<td width="100%" height="26" colspan="2" class=tablebody1>
	  说明：<li>若允许文件可删除，则当含有该文件的帖子被删除后，文件有可能亦被删除；
	  <li>文件说明不得超过１２５个汉字或２５０个字符．
	</td>
    	</tr>
	<tr>
	<th width="100%" height="26" colspan="2">
	<input type=Submit value="保存" name=Submit>&nbsp;
	</th>
	</tr>
	</form>
  </table>
<%
end if 
rs.close
set rs=nothing
end sub 


'保存修改
sub filesave()
if GroupSetting(48)=0 then
	errmsg=errmsg+"<br>"+"<li>您没有论坛文件管理的权限,请与管理员联系。"
	founderr=true
	exit sub
end if
dim saveid,F_Readme,Fflag,fileurl
dim F_BoardID,F_AnnounceID

F_BoardID=checkStr(trim(request.form("F_BoardID")))
F_AnnounceID=checkStr(trim(request.form("F_AnnounceID")))
F_Readme=checkStr(trim(request.form("F_Readme")))
saveid=trim(request.form("saveid"))
Fflag=trim(request.form("Fflag"))
if not isnumeric(Fflag) or not isnumeric(F_BoardID) then 
	errmsg=errmsg+"<br>"+"<li>执行的数据不存在,或不合符要求。"
	founderr=true
	exit sub
end if
if not isnumeric(saveid) or isnull(saveid) then 
	errmsg=errmsg+"<br>"+"<li>执行的数据不存在。"
	founderr=true
	exit sub
end if

if  F_Readme="" or isnull(F_Readme) then 
	errmsg=errmsg+"<br>"+"<li>说明内容不能为空。"
	founderr=true
	exit sub
end if

if strLength(F_Readme)>250 then
   		ErrMsg=ErrMsg+"<Br>"+"<li>说明内容不得大于" & 250 & "bytes"
   		FoundErr=true
		exit sub
end if
fileurl=checkStr(trim(request.form("fileurl")))
if fileurl<>"" then
	if instr(fileurl,"/")=0 or instr(fileurl,"://")=0  or instr(fileurl,".")=0 then
   		ErrMsg=ErrMsg+"<Br>"+"<li>请填写完整的文件链接地址！"
   		FoundErr=true
		exit sub
	end if
	conn.execute("update [DV_Upfile]  set F_Filename='"&fileurl&"',F_Readme='"&F_Readme&"' ,F_Flag="&Fflag&" where  F_ID="&saveid)
else
	conn.execute("update [DV_Upfile]  set F_BoardID="&F_BoardID&",F_AnnounceID='"&F_AnnounceID&"',F_Readme='"&F_Readme&"' ,F_Flag="&Fflag&" where  F_ID="&saveid)
end if
	sucmsg=sucmsg+"<br>"+"<li><b>编辑成功。<a href=myfile.asp >[返回管理页面]</a>"
	call dvbbs_suc()
end sub

'新增文件
sub addnew()
%>
 <table cellpadding=3 cellspacing=1 align=center class=tableborder1>
 <form action="?action=fsnew" method=post >
    <tr>
      <th width="100%" height="26" colspan="2" align=left>新增我个人文件</th>
    </tr>
	    <tr>
      <td width="15%" height="26" class=tablebody2 align=right>新增的文件：</td>
      <td width="85%"  height="26" class=tablebody1 align=left>
	  <input type=text name="fileurl" size="80%"  value="http://"> （请填写完成的文件链接地址）
	</td>
    </tr>
    <tr>
      <td width="15%" height="26" class=tablebody2 align=right>文件的说明：</td>
      <td width="85%"  height="26" class=tablebody1 align=left>
	  <textarea name="F_Readme" rows=6 cols=80%  wrap=VIRTUAL value=""></textarea> （不得超过１２５个汉字或２５０个字符）
	</td>
    </tr>
	<tr>
	 <td valign=middle align=right  class=tablebody2 >文件的类型：</td>
      <td valign=middle class=tablebody1>
	  <select name="filetype" size=1>
<option value="">文件类型</option>
<%
dim iupload
iupload="文件集|图片集|FLASH集|音乐集|电影集"
iupload=split(iupload,"|")
for i=0 to ubound(iupload)
response.write "<option value="&i&">"&iupload(i)&"</option>"
next
%>
</select>
	</td>
	</tr>
	 <td valign=middle align=right  class=tablebody2 >文件允许删除：</td>
      <td valign=middle class=tablebody1>
	  <input type=radio name=Fflag value=1  checked >是&nbsp;
	  <input type=radio name=Fflag value=2  >否
	</td>
	</tr>
	<tr>
		<td width="100%" height="26" colspan="2" class=tablebody1>
	  说明：<li>若允许文件可删除，当清理时会自动清除；
	  <li>文件说明不得超过１２５个汉字或２５０个字符；
	  <li>请用ＨＴＴＰ：／／或ＦＴＰ：／／或ＨＴＴＰＳ：／／开头的完整链接地址．
		</td>
    </tr>
    <tr>
		<th width="100%" height="26" colspan="2">
	  <input type=Submit value="保存" name=Submit>&nbsp;<INPUT TYPE="reset" value="清除">
		</th>
    </tr>
	</form>
  </table>
  <%
end sub 

'保存新增文件
sub savenew()
if GroupSetting(48)=0 then
	errmsg=errmsg+"<br>"+"<li>您没有论坛文件管理的权限,请与管理员联系。"
	founderr=true
	exit sub
end if
dim F_Readme,fileurl,filetype,fileExt,fileExt_a,filename
dim F_Type,F_Flag
F_Readme=checkStr(trim(request.form("F_Readme")))
fileurl=checkStr(trim(request.form("fileurl")))
filetype=trim(request.form("filetype"))
F_Flag=trim(request.form("Fflag"))

if  fileurl="" or isnull(fileurl) then 
	errmsg=errmsg+"<br>"+"<li>文件链接不能为空。"
	founderr=true
	exit sub
else
	if strLength(fileurl)>250 then
   		ErrMsg=ErrMsg+"<Br>"+"<li>文件链接不得大于" & 250 & "bytes"
   		FoundErr=true
		exit sub
	end if
end if

if  F_Readme="" or isnull(F_Readme) then 
	errmsg=errmsg+"<br>"+"<li>说明内容不能为空。"
	founderr=true
	exit sub
end if
if strLength(F_Readme)>250 then
   		ErrMsg=ErrMsg+"<Br>"+"<li>说明内容不得大于" & 250 & "bytes"
   		FoundErr=true
		exit sub
end if
if not isnumeric(F_Flag)  then 
	errmsg=errmsg+"<br>"+"<li>执行的数据不存在。"
	founderr=true
	exit sub
end if

	if instr(fileurl,"/")=0 or instr(fileurl,"://")=0  or instr(fileurl,".")=0 then
   		ErrMsg=ErrMsg+"<Br>"+"<li>请填写完整的文件链接地址！"
   		FoundErr=true
		exit sub
		else
		filename=split(fileurl,"/")
		fileExt=lcase(filename(ubound(filename)))
	end if

	fileExt_a=split(fileExt,".")
	fileExt=lcase(fileExt_a(ubound(fileExt_a)))
	
	if fileEXT="asp" and fileEXT="asa" and fileEXT="aspx" then
	   	ErrMsg=ErrMsg+"<Br>"+"<li>文件格式不支持，请重新填写！"
   		FoundErr=true
		exit sub
	end if
	
	if lcase(fileExt)="gif" or lcase(fileExt)="jpg" or lcase(fileExt)="jpeg" or lcase(fileExt)="bmp" or lcase(fileExt)="png" then
	F_Type=1
	elseif lcase(fileExt)="swf" or lcase(fileExt)="swi" then
	F_Type=2
	elseif lcase(fileExt)="mid" or lcase(fileExt)="wav" or lcase(fileExt)="mp3" or lcase(fileExt)="rmi" or lcase(fileExt)="cda" then
	F_Type=3
	elseif lcase(fileExt)="avi" or lcase(fileExt)="wov" or lcase(fileExt)="asf" or lcase(fileExt)="mpg" or lcase(fileExt)="mpeg" or lcase(fileExt)="ra" or lcase(fileExt)="ram" then
	F_Type=4
	else
	F_Type=0
	end if
	BoardID=0
conn.execute("insert into dv_upfile (F_BoardID,F_UserID,F_Username,F_Filename,F_FileType,F_Type,F_Readme,F_Flag ) values ("&BoardID&","&UserID&",'"&membername&"','"&replace(fileurl,"|","")&"','"&replace(fileExt,".","")&"',"&F_Type&",'"&F_Readme&"',"&F_Flag&" )")
	sucmsg=sucmsg+"<br>"+"<li><b>编辑成功。<a href=myfile.asp >[返回管理页面]</a>"
	call dvbbs_suc()
end sub

'删除文件
sub fdel()
dim delid
if GroupSetting(48)=0 then
	errmsg=errmsg+"<br>"+"<li>您没有论坛文件管理的权限,请与管理员联系。"
	founderr=true
	exit sub
end if
delid=replace(request("delid"),"'","")
if delid="" or isnull(delid) then
Errmsg=Errmsg+"<li>"+"请选择相关参数。"
founderr=true
exit sub
else
	if  master or superboardmaster or boardmaster  then
	conn.execute("delete from DV_Upfile where   F_ID in ("&delid&")")
	else
	conn.execute("delete from DV_Upfile where F_Flag=1 and  F_ID in ("&delid&")")
	end if
	sucmsg=sucmsg+"<br>"+"<li><b>您已经删除选定的文件记录。"
	call dvbbs_suc()
end if
end sub

'删除所有文件
sub alldel()
if GroupSetting(48)=0 then
	errmsg=errmsg+"<br>"+"<li>您没有论坛文件管理的权限,请与管理员联系。"
	founderr=true
	exit sub
end if
dim delid
delid=replace(request("delid"),"'","")
if delid="" or isnull(delid) then
Errmsg=Errmsg+"<li>"+"请选择相关参数。"
founderr=true
exit sub
end if
	'conn.execute("delete from DV_Upfile where F_Flag=1 and  F_UserID="&userid)
	sucmsg=sucmsg+"<br>"+"<li><b>您已经删除了所有文件记录。"
	call dvbbs_suc()
end sub 

function filenum(types)
dim stype
if isnumeric(types)  then
	select case types
	case 0
	stype="F_Type=0 and"
	case 1
	stype="F_Type=1 and"
	case 2
	stype="F_Type=2 and"
	case 3
	stype="F_Type=3 and"
	case 4
	stype="F_Type=4 and"
	case else
	stype=""
	end select
else
	stype=""
end if
	rs=conn.execute("Select Count(F_ID) From DV_Upfile Where  "&stype&" F_UserID="&userid&"")
	filenum=rs(0)
	set rs=nothing
	if isnull(filenum) then filenum=0
end function

%>


