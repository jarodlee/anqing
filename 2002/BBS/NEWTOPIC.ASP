<!--#include file="conn.asp"-->
<%  

dim bbsurl
bbsurl="http://www.szzs.com/bbs/"  '请填写你论坛的正确地址,要以"HTTP://"开头

'bbsurl=getservepath(request.ServerVariables("server_name")&request.ServerVariables("URL"))
'function getservepath(str)
'dim tmpstr
'tmpstr=split(str,"/")
'getservepath="http://"&replace(str, tmpstr(ubound(tmpstr)), "")
'end function


'*************************************
'上传到与CONN.ASP同级的目录下
'以上地址参数一定要修改,否则所调用的链接是去了以上的论坛.
'若有问题,可以运行一起上传的newscode.ASP文件进行调试(newscode.ASP运行前要修改调用参数)
'	FSSUNWIN	2002.10.18
'*************************************

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

    HTMLEncode = fString
end if
end function

	dim rs,sql
	dim orders,reply,topic,isbest
        dim board
	if request("orders")=1 then
		orders="hits"
	else
		orders="dateandtime"
	end if
        if request("boardid")="all" then
            board=""
        elseif isnumeric(request("boardid")) then
            board=" and boardid="&request("boardid")&""
	else
	    board=" and boardid=1"
        end if
	if request("action")=1 then
	'显示主题
          set rs=conn.execute("select top "&replace(request("n"),"'","")&" PostUserName,Title,topicid,boardid,dateandtime,topicid,hits from topic where  datediff('d',dateandtime,Now())>="&replace(request("sdate"),"'","")-1&" and not locktopic=2 "&board&"  ORDER BY "&orders&" desc")
	elseif request("action")=2 then
        '显示精华主题
          set rs=conn.execute("select top "&replace(request("n"),"'","")&" PostUserName,Title,rootid,boardid,dateandtime,Announceid,id from BestTopic where  datediff('d',dateandtime,Now())>="&replace(request("sdate"),"'","")-1&"  "&board&"  ORDER BY "&orders&" desc")
        else
          set rs=conn.execute("select top "&replace(request("n"),"'","")&" username,topic,rootid,boardid,dateandtime,announceid,body from bbs1 where  datediff('d',dateandtime,Now())>="&replace(request("sdate"),"'","")-1&" and not locktopic=2 "&board&" ORDER BY "&orders&" desc")
        '显示主题或回复
        end if

	do while Not RS.Eof
       topic=rs(1)
       if topic="" then
          topic=rs(6)
       end if
       if len(topic)>Cint(request("tlen")) then
          topic=left(topic,request("tlen"))&"..."
          end if
       topic=htmlencode(topic)
	response.write "document.write('<FONT color=#b70000><B>·</B></FONT><span style=""font-size:9pt;line-height: 15pt""><a href="&bbsurl&"dispbbs.asp?boardid="&rs(3)&"&ID="&rs(2)&"&replyID="&rs(5)&" target=_top title="&htmlencode(rs(1))&">');"
        response.write "document.write('"&topic&"');"
	response.write "document.write('</a>');"
	select case request("info")
	case 0
	case 1
	response.write "document.write('(<a href="&bbsurl&"dispuser.asp?name="&rs(0)&" target=_blank>"&rs(0)&"</a>,<font color=green>"&formatdatetime(rs("dateandtime"),0)&"</font>)');"
	case 2
	response.write "document.write('(<font color=green>"&rs("dateandtime")&"</font>)');"
	case 3
	response.write "document.write('(<a href="&bbsurl&"dispuser.asp?name="&rs(0)&" target=_blank>"&rs(0)&"</a>)');"
	case 4
	response.write "document.write('(<a href="&bbsurl&"dispuser.asp?name="&rs(0)&" target=_blank>"&rs(0)&"</a>,<font color=green>"&rs(6)&"</font>)');"
	case 5
	response.write "document.write('(<font color=green>"&rs(6)&"</font>)');"
	case 6
	response.write "document.write('(<a href="&bbsurl&"dispuser.asp?name="&rs(0)&" target=_blank>"&rs(0)&"</a>,<font color=green>"&formatdatetime(rs("dateandtime"),1)&"</font>)');"
	case 7
	response.write "document.write('(<font color=green>"&formatdatetime(rs("dateandtime"),1)&"</font>)');"
	case else

	end select

	response.write "document.write('</span><br>');"
	RS.MoveNext
	Loop
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
%>

