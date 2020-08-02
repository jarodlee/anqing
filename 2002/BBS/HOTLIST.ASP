<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
	'on error resume next
	dim stype,ordername

	dim currentpage,page_count,Pcount
	dim totalrec,endpage

	dim searchDateLimit,searchday
	dim searchboard

	dim LastPost,Lastuser,LastID
	dim LastTime,LastUserid,LastRootid,body

	stats="热门话题"

	stype=request("stype")
	if request("page")="" or not isInteger(request("page"))  then
		currentPage=1
	else
		currentPage=cint(request("page"))
	end if
	'搜索多少天内帖子
	searchDateLimit=180
	if request("SearchDate")="ALL" OR request("SearchDate")="" then
	searchday=" "
	else
		if not isInteger(request("SearchDate"))  then
			Errmsg=Errmsg+"<br>"+"<li>搜索多少天必须是整数。"
			founderr=true
		else
			searchday=" datediff('d',dateandtime,Now()) < "&request("SearchDate")&" and "
		end if
	end if


	call nav()
	if cint(boardid)<1 then
	call head_var(2,0,"","hotlist.asp?boardid=0")
	searchboard=""
	else
	call head_var(1,0,boardtype,"hotlist.asp?boardid="&boardid)
	searchboard=" boardid="&boardid&" and "
	end if

	if Cint(GroupSetting(14))=0 then
		Errmsg=Errmsg+"<br>"+"<li>您没有在本论坛搜索的权限，请<a href=login.asp>登陆</a>或者同管理员联系。"
		founderr=true
	end if
	if founderr then
		call dvbbs_error()
	else
		call search()
		if founderr then call dvbbs_error()
	end if
	call footer()

	sub search()
	sql=""
	set rs=server.createobject("adodb.recordset")
	select case stype
	case 1
		sql="select top 50  TopicID,boardid,title,postusername,postuserid,dateandtime,child,hits,votetotal,lastpost,lastposttime,isvote,isbest,locktopic,Expression from Topic where "&searchboard&" "&searchday&"   locktopic<2 and hits>10 ORDER BY hits desc"
		ordername="最热门50主题"

	case 2
		sql="select top 50  TopicID,boardid,title,postusername,postuserid,dateandtime,child,hits,votetotal,lastpost,lastposttime,isvote,isbest,locktopic,Expression from Topic where "&searchboard&" "&searchday&"   locktopic<2 ORDER BY hits desc"
		ordername="最新50帖"
	case else
		sql="select top 50  TopicID,boardid,title,postusername,postuserid,dateandtime,child,hits,votetotal,lastpost,lastposttime,isvote,isbest,locktopic,Expression from Topic where "&searchboard&" "&searchday&"   locktopic<2 and hits>10 ORDER BY hits desc"
		ordername="最热门50主题"

	end select
	
	if sql="" then
		Errmsg=Errmsg+"<br>"+"<li>请指定查询条件。"
		founderr=true
		exit sub
	end if
	rs.open sql,conn,1,1

	if rs.eof and rs.bof then
		Errmsg=Errmsg+"<br>"+"<li>没有找到您要查询的内容。"
		founderr=true
		exit sub
	else
		rs.PageSize = Forum_Setting(11)
		rs.AbsolutePage=currentpage
		page_count=0
		totalrec=rs.recordcount
		call searchinfo()
		call listPages3()
	end if

	rs.close
	set rs=nothing
	call activeonline()
	end sub

	sub searchinfo()

%>
	<form action="<%=ScriptName%>" method=get><table cellpadding=0 cellspacing=0 border=0 width="<%=Forum_body(12)%>" align=center>
	<tr><td>目前共有<font color=<%=Forum_body(8)%>><%=totalrec%></font>个<%=ordername%>
	</td>
	<td width=200 align=right>
	<input type=hidden name=boardid value="<%= boardid %>">
	<input type=hidden name=stype value="<%=request("stype")%>">
	<select name=SearchDate onchange='javascript:submit()'>
	<option value=all>查看所有的主题
	<option value=1>查看一天内的主题
	<option value=2>查看两天内的主题
	<option value=7>查看一星期内的主题
	<option value=15>查看半个月内的主题
	<option value=30>查看一个月内的主题
	<option value=60>查看两个月内的主题
	<option value=180>查看半年内的主题
	</select>
	</td></tr>
	</table></form>
<TABLE cellPadding=3 cellSpacing=1 class=tableborder1 align=center>
<TR valign=middle>
<Th height=25 width=32>状态</Th>
<Th width=*>主 题</Th>
<Th width=80>作 者</Th>
<Th width=64>回复/人气</Th>
<Th width=195>最后更新 | 回复人</Th>
</TR>
<%
       while (not rs.eof) and (not page_count = rs.PageSize)

if not isnull(rs("lastpost")) then
	LastPost=split(rs("lastpost"),"$")
	if ubound(LastPost)=6 then
	Lastuser=htmlencode(LastPost(0))
	LastID=LastPost(1)
	LastTime=LastPost(2)
	if not isdate(LastTime) then LastTime=rs("dateandtime")
	body=htmlencode(LastPost(3))
	LastUserid=LastPost(5)
	LastRootid=LastPost(6)
	else
	Lastuser=htmlencode(LastPost(0))
	LastID=rs("topicid")
	LastTime=rs("dateandtime")
	body="..."
	LastUserid=rs("postuserid")
	LastRootid=rs("topicid")
	end if
else
	Lastuser="------"
	LastID=rs("topicid")
	LastTime=rs("dateandtime")
	body="..."
	LastUserid=rs("postuserid")
	LastRootid=rs("topicid")
end if

response.write"<TR><TD align=middle class=tablebody2 width=32>"
if rs(0)=1 then
response.write"<img src="&Forum_info(7)&Forum_statePic(2)&" alt=""本主题已锁定"">"
else
response.write"<img src="&Forum_info(7)&Forum_statePic(1)&" >"
end if
response.write"</TD><TD class=tablebody1 width=*><a href='dispbbs.asp?boardID="&rs("boardID")&"&ID="& rs("topicid")&"&skin=1' target=_blank>"
response.write"<img src="&Forum_info(8)&""
if instr(rs("Expression"),Forum_Setting(44))>0 then
response.write rs("Expression")
else
response.write "face1.gif"
end if
response.write  " border=0 alt=""开新窗口浏览此主题""></a>"&_ 
				"<a href='dispbbs.asp?boardID="&rs("boardID")&"&ID="& rs("topicid") &"&skin=1'>"
if renzhen(rs("boardid"),membername)=false then
	response.write "<font color=gray>（认证论坛帖子，只有认证用户才能查看）</font>"
else
	if len(rs("title"))>26 then
		response.write ""&left(htmlencode(replace(rs("title"),chr(10)," ")),26)&"..."
	else
		response.write htmlencode(rs("title"))
	end if
end if
response.write "</a></TD>"&_
				"<TD align=middle  class=tablebody2  width=80><a href=""dispuser.asp?id="&rs("postuserid")&""">"&htmlencode(rs("postusername"))&"</a></TD>" 

response.write "<TD class=tablebody1 width=64 align=center>"
if rs("isvote")=1 then
	response.write "<FONT color="&Forum_body(8)&"><b>"&rs("votetotal")&"</b></font>  票"
else
	response.write rs("child") &"/"& rs("hits")
end if
response.write "</TD>" 
%>
<TD  class=tablebody2 width=195>&nbsp;<a href='dispbbs.asp?boardid=<%=rs("boardid")%>&id=<%=LastRootID%>&<%=LastID%>' >
<%=FormatDateTime(LastTime,2)%>&nbsp;<%=FormatDateTime(LastTime,4)%></a>
&nbsp;<font color="<%=Forum_body(8)%>">|</font>&nbsp;
<a href="dispuser.asp?id=<%=LastUserID%>" target=_blank><%=htmlencode(LastUser)%></a>
</TD>
</TR> 
<%
	page_count = page_count + 1
          rs.movenext
        wend
	response.write "</table>"

	end sub

	sub listPages3()

	Pcount=rs.PageCount
	response.write "<table border=0 cellpadding=0 cellspacing=3 width="""&Forum_body(12)&""" align=center>"&_
			"<tr><td valign=middle nowrap>"&_
			"<font color="&Forum_body(13)&">页次：<b>"&currentpage&"</b>/<b>"&Pcount&"</b>页"&_
			"每页<b>"&Forum_Setting(11)&"</b> 帖子数<b>"&totalrec&"</b></font></td>"&_
			"<td valign=middle nowrap><font color="&Forum_body(13)&"><div align=right><p>分页： "

	if currentpage > 4 then
	response.write "<a href=""?page=1&stype="&stype&"&SearchDate="&request("SearchDate")&"&boardid="&boardid&""">[1]</a> ..."
	end if
	if Pcount>currentpage+3 then
	endpage=currentpage+3
	else
	endpage=Pcount
	end if
	for i=currentpage-3 to endpage
	if not i<1 then
		if i = clng(currentpage) then
        response.write " <font color="&Forum_body(8)&">["&i&"]</font>"
		else
        response.write " <a href=""?page="&i&"&stype="&stype&"&SearchDate="&request("SearchDate")&"&boardid="&boardid&""">["&i&"]</a>"
		end if
	end if
	next
	if currentpage+3 < Pcount then 
	response.write "... <a href=""?page="&Pcount&"&stype="&stype&"&SearchDate="&request("SearchDate")&"&boardid="&boardid&""">["&Pcount&"]</a>"
	end if
	response.write "</p></div></font></td></tr></table>"

	end sub
function renzhen(boardid,username)
	dim boarduser,rrs,Board_Setting,BoardMaster
	renzhen=false
	if master then
		renzhen=true
	else
		sql="select boarduser,Board_Setting,BoardMaster from board where boardid="&boardid
		set rrs=server.createobject("adodb.recordset")
		rrs.open sql,conn,1,1
		Board_Setting=split(rrs("board_setting"),",")
		if cint(Board_Setting(2))=1 then
			if not (isnull(rrs(2)) or rrs(2)="") then
				BoardMaster=split(rrs(2), "|")
				for i = 0 to ubound(BoardMaster)
					if trim(BoardMaster(i))=trim(username) then
						renzhen=true
						exit for
					end if
				next
			end if
			if renzhen=false then
				if isnull(rrs(0)) or rrs(0)="" then
					renzhen=false
				else
					boarduser=split(rrs(0), ",")
					for i = 0 to ubound(boarduser)
						if trim(boarduser(i))=trim(username) then
							renzhen=true
							exit for
						end if
					next
				end if
			end if
		else
			renzhen=true
		end if
		rrs.close
		set rrs=nothing		
	end if
end function
%>