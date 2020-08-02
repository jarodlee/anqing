<!--#include file="conn.asp"-->
<% 
    dim rs,sql,TopicNum,BbsNum,TodayNum,UserNum,lastUser,tmprs
    dim n
    dim maxpostnum,yesterdaynum
    if request("orders")=1 then
       call tongji()
    elseif request("orders")=2 then
       call topuser()
    else
       call tongji()
       end if



function online() 
     tmprs=conn.execute("Select count(id) from online")  
     online=tmprs(0)  
 set tmprs=nothing  
 if isnull(online) then online=0 
end function
            
function allboys()  
dim tmprs  
    tmprs=conn.execute("Select count(userid) from [user] where sex='1' ")  
    allboys=tmprs(0)  
 set tmprs=nothing  
 if isnull(allboys) then allboys=0  
end function

function allgirls()  
dim tmprs  
    tmprs=conn.execute("Select count(userid) from [user] where sex='0' ")  
    allgirls=tmprs(0)  
 set tmprs=nothing  
 if isnull(allgirls) then allgirls=0  
end function

 sub tongji()
   n=request("n")
   

sql="select top 1 TopicNum,BbsNum,TodayNum,UserNum,lastUser,yesterdaynum,maxpostnum,maxpostdate from config where active=1"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
        topicnum=rs(0)
        BbsNum=rs(1)
        TodayNum=rs(2)
        UserNum=rs(3)
        lastUser=rs(4)
        yesterdaynum=rs(5)
        maxpostnum=rs(6)
       rs.close
   set rs=nothing
%> 
 document.write('□- 论坛贴数 <font color=green><%=BbsNum%></font><br> □- 论坛主题 <font color=green><%=TopicNum%></font><br> □- 注册人数 <font color=green><%=UserNum%></font><br> □- 论坛在线 <font color=red><%=online()%></font>');
 document.write('<br> □- 今日帖数 <font color=red><%=TodayNum%></font>')

<% end sub
sub topuser()
	set rs=server.createobject("adodb.recordset")
	sql="select top "&request("n")&" username,useremail,userclass,oicq,homepage,article,addDate  from [user] order by article desc"
	rs.open sql,conn,1,1
	do while Not RS.Eof

	response.write "document.write('<FONT color=#b70000>□</FONT><span style=""font-size:9pt;line-height: 15pt""><a href=http://www.pydz.com/sunwinbbs/dispuser.asp?name="& rs(0) &" target=_blank title=查看"&rs(0)&"的个人资料>');"
        response.write "document.write('"&rs(0)&"</a>');"
	response.write "document.write('</span><br>');"

	rs.MoveNext
	Loop
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing

end sub
%>