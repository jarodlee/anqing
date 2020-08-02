<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
	'on error resume next
	dim currentpage,page_count,Pcount
	dim totalrec,endpage
	dim stype,pSearch,nSearch,keyword
	dim searchboard,ordername
	dim searchdate,searchDateLimit,searchday
	dim stable
	stats="搜索结果"
	stype=request("stype")
	pSearch=request("pSearch")
	nSearch=request("nSearch")
	keyword=trim(checkStr(request("keyword")))
	stable=checkstr(request("stable"))
	if stable="" then stable=Nowusebbs
	if request("page")="" or not isInteger(request("page"))  then
		currentPage=1
	else
		currentPage=cint(request("page"))
	end if
	if stype<3 then
	if keyword="" then
		Errmsg=Errmsg+"<br>"+"<li>必须输入查询关键字。"
		founderr=true
	end if

	'搜索多少天内帖子
	searchDateLimit=180
	if request("SearchDate")="ALL" then
	searchday=" "
	else
		if not isInteger(request("SearchDate"))  then
			Errmsg=Errmsg+"<br>"+"<li>搜索多少天必须是整数。"
			founderr=true
		else
			searchday=" datediff('d',dateandtime,Now()) < "&request("SearchDate")&" and "
		end if
	end if
	end if

	call nav()
	if cint(boardid)<1 then
	call head_var(0,0,"论坛搜索","query.asp?boardid=0")
	searchboard=""
	else
	call head_var(0,0,boardtype,"query.asp?boardid="&boardid)
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
		select case nSearch
		'搜索主题帖子作者
		case 1
		sql="select top 1000 locktopic,boardid,rootid,announceid,body,Expression,topic,username,dateandtime,postuserid,ParentID from "&stable&" where username='"&keyword&"' and "&searchboard&" "&searchday&" parentid=0 and locktopic<2 ORDER BY announceID desc"
		ordername="搜索主题作者帖子"
		'搜索回复帖子作者
		case 2
		sql="select top 1000 locktopic,boardid,rootid,announceid,body,Expression,topic,username,dateandtime,postuserid,ParentID from "&stable&" where username='"&keyword&"' and "&searchboard&" "&searchday&" parentid>0 and locktopic<2 ORDER BY announceID desc"
		ordername="搜索回复作者帖子"
		'两者都搜索
		case 3
		sql="select top 1000  locktopic,boardid,rootid,announceid,body,Expression,topic,username,dateandtime,postuserid,ParentID from "&stable&" where  "&searchboard&" "&searchday&" username='"&keyword&"' and locktopic<2 ORDER BY announceID desc"
		ordername="搜索主题和回复作者帖子"
		end select
	case 2
		select case pSearch
		'搜索主题关键字
		case 1
		sql="select top 1000 locktopic,boardid,rootid,announceid,body,Expression,topic,username,dateandtime,postuserid,ParentID from "&stable&" where  "&searchboard&" "&searchday&" (" & translate(keyword,"topic") & ") and locktopic<2 ORDER BY announceID desc"
		'搜索内容关键字
		ordername="搜索主题"
		case 2
		sql="select top 1000  locktopic,boardid,rootid,announceid,body,Expression,topic,username,dateandtime,postuserid,ParentID from "&stable&" where   "&searchboard&" "&searchday&" (" & translate(keyword,"body") & ") and locktopic<2 ORDER BY announceID desc"
		'两者都搜索
		ordername="搜索内容"
		case 3
		sql="select top 1000  locktopic,boardid,rootid,announceid,body,Expression,topic,username,dateandtime,postuserid,ParentID from "&stable&" where   "&searchboard&" "&searchday&" (" & translate(keyword,"topic") & " or " & translate(keyword,"body") & ") and locktopic<2 ORDER BY announceID desc"
		ordername="搜索主题和内容"
		end select
	case 3
		sql="select top 50  locktopic,boardid,rootid,announceid,body,Expression,topic,username,dateandtime,postuserid,ParentID from "&stable&" where locktopic<2 ORDER BY announceID desc"
		ordername="最新50帖"
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
            <table cellpadding=0 cellspacing=0 border=0 width="<%=Forum_body(12)%>" align=center>
            <tr><td>查询<%if request("SearchDate")="ALL" then%>所有日期<%else%><%=request("SearchDate")%>天内<%end if%>的帖子，<%=ordername%>
<%if totalrec>=1000 then%>
最多只能查询到<font color=<%=Forum_body(8)%>>1000</font>条纪录，更多的就不显示了
<%else%>共查询到
<font color=<%=Forum_body(8)%>><%=totalrec%></font>个结果
<%end if%>
		</td>
            </tr>
            </table>
<TABLE cellPadding=3 cellSpacing=1 class=tableborder1 align=center>
<TR valign=middle>
<Th height=25 width=32>状态</Th>
<Th width=*>主 题</Th>
<Th width=80>作 者</Th>
<Th width=195>最后更新 | 回复人</Th>
</TR>
<%
       while (not rs.eof) and (not page_count = rs.PageSize)
%>
  <TR> 
    <TD align=middle class=tablebody2 width=32>
<%
'帖子状态修改 杨铮20030108
dim RsTopic,lock
lock=0
if rs(10)=0 then
	set RsTopic=conn.execute("select * from topic where topicid="&rs(2))
	if RsTopic("istop")=2 then lock=1
	if RsTopic("istop")=1 then lock=2
	if RsTopic("isbest")=1 then lock=3
	if RsTopic("isvote")=1 then lock=4
	if RsTopic("locktopic")=1 then lock=5
	if RsTopic("child")>=Cint(forum_setting(44)) then lock=6
end if
if lock=1 then%>
	<img src=<%=Forum_info(7)&Forum_statePic(9)%> alt=总固顶主题>
<%elseif lock=2 then%>
	<img src=<%=Forum_info(7)&Forum_statePic(3)%> alt=固顶主题>
<%elseif lock=3 then%>
	<img src=<%=Forum_info(7)&Forum_statePic(4)%> alt=精华帖子>
<%elseif lock=4 then%>
	<img src=<%=Forum_info(7)&Forum_statePic(12)%> alt=投票贴子>
<%elseif lock=5 then%>
	<img src=<%=Forum_info(7)&Forum_statePic(2)%> alt=本主题已锁定>
<%elseif lock=6 then%>
	<img src=<%=Forum_info(7)&Forum_statePic(1)%> alt=热门主题>
<%else%>
	<img src=<%=Forum_info(7)&Forum_statePic(0)%> alt=开放主题或回帖>
<%end if%>
    </TD>
    <TD  class=tablebody1 width=*><a href='dispbbs.asp?boardID=<%=rs(1)%>&replyID=<%=rs(3)%>&ID=<%=rs(2)%>&skin=1' target=_blank><img src='<%=Forum_info(8)&rs(5)%>' border=0 alt="开新窗口浏览此主题"></a> <a href='dispbbs.asp?boardID=<%=rs(1)%>&replyID=<%=rs(3)%>&ID=<%=rs(2)%>&skin=1'>
<%
'加入认证帖子判断 杨铮20030108
if renzhen(rs(1),membername)=false then%>
	<font color=gray>（认证论坛帖子，只有认证用户才能查看）</font>
<%else
if rs(6)="" then%>
	<%=left(htmlencode(replace(rs(4),chr(10)," ")),26)%>...
<%else%>
	<%=htmlencode(rs(6))%>
<%end if
end if%>
</a>    </TD> 
    <TD align=middle  class=tablebody2  width=80><a href="dispuser.asp?id=<%=htmlencode(rs("postuserid"))%>"><%=htmlencode(rs(7))%></a></TD> 
    <TD  class=tablebody1 width=195>&nbsp;
<%=FormatDateTime(rs("dateandtime"),2)%>&nbsp;<%=FormatDateTime(rs("dateandtime"),4)%>
&nbsp;<font color="<%=Forum_body(8)%>">|</font>&nbsp;
<a href="dispuser.asp?id=<%=rs("postuserid")%>" target=_blank><%=htmlencode(rs(7))%></a>
</TD>
</TR> 
<%
	  page_count = page_count + 1
          rs.movenext
        wend
%>
            </table>
<%
	end sub

	sub listPages3()

	Pcount=rs.PageCount
	response.write "<table border=0 cellpadding=0 cellspacing=3 width="""&Forum_body(12)&""" align=center>"&_
			"<tr><td valign=middle nowrap>"&_
			"<font color="&Forum_body(13)&">页次：<b>"&currentpage&"</b>/<b>"&Pcount&"</b>页"&_
			"每页<b>"&Forum_Setting(11)&"</b> 帖子数<b>"&totalrec&"</b></font></td>"&_
			"<td valign=middle nowrap><font color="&Forum_body(13)&"><div align=right><p>分页： "

	if currentpage > 4 then
	response.write "<a href=""?page=1&stype="&stype&"&pSearch="&pSearch&"&nSearch="&nSearch&"&keyword="&keyword&"&SearchDate="&request("SearchDate")&"&boardid="&boardid&"&stable="&stable&""">[1]</a> ..."
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
        response.write " <a href=""?page="&i&"&stype="&stype&"&pSearch="&pSearch&"&nSearch="&nSearch&"&keyword="&keyword&"&SearchDate="&request("SearchDate")&"&boardid="&boardid&"&stable="&stable&""">["&i&"]</a>"
		end if
	end if
	next
	if currentpage+3 < Pcount then 
	response.write "... <a href=""?page="&Pcount&"&stype="&stype&"&pSearch="&pSearch&"&nSearch="&nSearch&"&keyword="&keyword&"&SearchDate="&request("SearchDate")&"&boardid="&boardid&"&stable="&stable&""">["&Pcount&"]</a>"
	end if
	response.write "</p></div></font></td></tr></table>"

	end sub

public function translate(sourceStr,fieldStr)
rem 处理逻辑表达式的转化问题
  dim  sourceList
  dim resultStr
  dim i,j
  if instr(sourceStr," ")>0 then 
     dim isOperator
     isOperator = true
     sourceList=split(sourceStr)
     '--------------------------------------------------------
     rem Response.Write "num:" & cstr(ubound(sourceList)) & "<br>"
     for i = 0 to ubound(sourceList)
        rem Response.Write i 
    Select Case ucase(sourceList(i))
    Case "AND","&","和","与"
        resultStr=resultStr & " and "
        isOperator = true
    Case "OR","|","或"
        resultStr=resultStr & " or "
        isOperator = true
    Case "NOT","!","非","！","！"
        resultStr=resultStr & " not "
        isOperator = true
    Case "(","（","（"
        resultStr=resultStr & " ( "
        isOperator = true
    Case ")","）","）"
        resultStr=resultStr & " ) "
        isOperator = true
    Case Else
        if sourceList(i)<>"" then
            if not isOperator then resultStr=resultStr & " and "
            if inStr(sourceList(i),"%") > 0 then
                resultStr=resultStr&" "&fieldStr& " like '" & replace(sourceList(i),"'","''") & "' "
            else
                resultStr=resultStr&" "&fieldStr& " like '%" & replace(sourceList(i),"'","''") & "%' "
            end if
                isOperator=false
        End if    
    End Select
        rem Response.write resultStr+"<br>"
     next 
     translate=resultStr
  else '单条件
     if inStr(sourcestr,"%") > 0 then
         translate=" " & fieldStr & " like '" & replace(sourceStr,"'","''") &"' "
     else
    translate=" " & fieldStr & " like '%" & replace(sourceStr,"'","''") &"%' "
     End if
     rem 前后各加一个空格，免得连sql时忘了加，而出错。
  end if  
end function
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