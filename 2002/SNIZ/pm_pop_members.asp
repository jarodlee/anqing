<%
'########## Snitz Forums 2000 Version 3.1 SR4 ####################
'#                                                               #
'#  �����޸�: ��Դ����վ                                         #
'#  �����ʼ�: cgier@21cn.com                                     #
'#  ��ҳ��ַ: http://www.sdsea.com                               #
'#            http://www.99ss.net                                #
'#            http://www.cdown.net                               #
'#	     http://www.wzdown.com                               #
'#	     http://www.13888.net                                #
'#  ��̳��ַ:http://ubb.yesky.net                                #
'#  ����޸�����: 2001/03/12    ���İ汾��Version 3.1 SR4        #
'#################################################################
'# ԭʼ��Դ                                                      #
'# Snitz Forums 2000 Version 3.1 SR4                             #
'# Copyright 2000 http://forum.snitz.com - All Rights Reserved   #
'#################################################################
'#����Ȩ������                                                   #
'#                                                               #
'# ������Ϊ��������(shareware)�ṩ������վ���ʹ�ã�����Ƿ��޸�,#
'# ת�أ�ɢ��������������ͼ����Ϊ��������ɾ����Ȩ������          #
'# ���������վ��ʽ����������ű�������֪ͨ���ǣ��Ա������ܹ�֪��#
'# ������ܣ�����������վ�������ǵ����ӣ�ϣ���ܸ��������лл��  #
'#################################################################
'# �����������ǵ��Ͷ��Ͱ�Ȩ����Ҫɾ�����ϵİ�Ȩ�������֣�лл����#
'# �����κ������뵽���ǵ���̳��������                            #
'#################################################################
%>
<!--#INCLUDE FILE="config.asp" -->
<!--#INCLUDE FILE="inc_functions.asp" -->
<!--#INCLUDE FILE="inc_top_short.asp" -->

<%

mypage = request("whichpage")
if mypage = "" then
	mypage = 1
end if
mypagesize = request("pagesize")
if mypagesize = "" then
	mypagesize = 30
end if
mySQL = request("strSql")
if mySQL = "" then
	mySQL = SQLtemp
end if

If Request.QueryString("mode") = "search" then 
strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.MEMBER_ID " 
strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME LIKE '" & Request("M_NAME") & "%' "

Set rsMembers = Server.CreateObject("ADODB.RecordSet")
rsMembers.open  strSql, my_conn, 3

if not (rsMembers.EOF or rsMembers.BOF) then  '## No categories found in DB
	rsMembers.movefirst
	rsMembers.pagesize = mypagesize

	maxpages = cint(rsMembers.pagecount)
end if

%>
<table width="95%" border="0">
  <tr>
    <td align="right">
<% if maxpages > 1 then %>
    <table border=0 align="right">
      <tr>
        <td valign="top"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>ҳ����</b> &nbsp;&nbsp;</font></td>
        <td valign="top"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% Call Paging() %></font></td>
      </tr>
    </table>
    
<% else %>
    &nbsp;
<% end if %>
    </td>
  </tr>
</table>

<table border="0" width="95%" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td bgcolor="<% =strTableBorderColor %>">
    <table border="0" width="100%" cellspacing="1" cellpadding="4">
      <tr>
        <td align="center" bgcolor="<% =strHeadCellColor %>"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">��Ա�б�</font></b></td>

<% If rsMembers.EOF or rsMembers.BOF then  '## No Members Found in DB %>
      <tr>
        <td align="center" bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>Ŀǰû���κλ�Ա</b></font>
        <p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript:history.go(-1)">����</a></font></p>
        </td>
      </tr>
<% Else 
	currMember = 0 %>
<%
	i = 0
	rsMembers.cacheSize = 30
	rsMembers.moveFirst
	rsMembers.pageSize = myPageSize
	maxPages = cint(rsMembers.pageCount)
	maxRecs = cint(rsMembers.pageSize)
	rsMembers.absolutePage = myPage
	howManyRecs = 0
	rec = 1
	do until rsMembers.Eof or rec = 31 
		if i = 1 then 
			CColor = strAltForumCellColor
		else
			CColor = strForumCellColor
		end if
		
%>
      <tr>

        <td bgcolor="<% =CColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="pop_profile.asp?mode=display&id=<% =rsMembers("MEMBER_ID") %>" target="_new"><% =ChkString(rsMembers("M_NAME"),"display") %></a>&nbsp;&nbsp;<img src="<%=strImageURL %>pm.gif" width="11" height="17" alt="�������Ļ�ѶϢ" border="0"  style="cursor:hand" onClick="opener.document.PostTopic.sendto.value+='<% =ChkString(rsMembers("M_NAME"),"display") %>'"></font></td>

      </tr>
<%
		currMember = rsMembers("MEMBER_ID")
		rsMembers.MoveNext
		i = i + 1
		if i = 2 then i = 0
		rec = rec + 1
	loop %>
      <tr>
        <td align="center" bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript:history.go(-1)">����</a></font></td>
      </tr>
<% end if 
%>


      </table>
      </td></tr></table>
<!--#INCLUDE FILE="inc_footer_short.asp" -->
<% else
'## Forum_SQL - Get all active topics from last visit
strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_STATUS, " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.M_FIRSTNAME, " & strMemberTablePrefix & "MEMBERS.M_LASTNAME, " & strMemberTablePrefix & "MEMBERS.M_LEVEL, " & strMemberTablePrefix & "MEMBERS.M_EMAIL, " & strMemberTablePrefix & "MEMBERS.M_COUNTRY, " & strMemberTablePrefix & "MEMBERS.M_HOMEPAGE, " & strMemberTablePrefix & "MEMBERS.M_ICQ, " & strMemberTablePrefix & "MEMBERS.M_YAHOO, " & strMemberTablePrefix & "MEMBERS.M_AIM, " & strMemberTablePrefix & "MEMBERS.M_TITLE, " & strMemberTablePrefix & "MEMBERS.M_POSTS, " & strMemberTablePrefix & "MEMBERS.M_LASTPOSTDATE, " & strMemberTablePrefix & "MEMBERS.M_LASTHEREDATE, " & strMemberTablePrefix & "MEMBERS.M_DATE "
strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
if mlev = 4 then
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME <> 'n/a' "
else
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_STATUS = " & 1
end if


Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open  strSql, my_conn, 3

if not (rs.EOF or rs.BOF) then  '## No categories found in DB
	rs.movefirst
	rs.pagesize = mypagesize

	maxpages = cint(rs.pagecount)
end if
%>
<table width="95%" border="0">
  <tr>
    <td align="right">
<% if maxpages > 1 then %>
    <table border=0 align="right">
      <tr>
        <td valign="top"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>ҳ����</b> &nbsp;&nbsp;</font></td>
        <td valign="top"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% Call Paging() %></font></td>
      </tr>
    </table>
<% else %>
    &nbsp;
<% end if %>
    </td>
  </tr>
</table>

<table border="0" width="95%" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td bgcolor="<% =strTableBorderColor %>">
    <table border="0" width="100%" cellspacing="1" cellpadding="4">
      <tr>
        <td align="center" bgcolor="<% =strHeadCellColor %>"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">��Ա�б�</font></b></td>

<% If rs.EOF or rs.BOF then  '## No Members Found in DB %>
      <tr>
        <td align="center" bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>Ŀǰû���κλ�Ա</b></font>
        <p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript:history.go(-1)">����</a></font></p>
        </td>
      </tr>
<% Else %>
<%	currMember = 0 %>
<%
	i = 0
	rs.cacheSize = 30
	rs.moveFirst
	rs.pageSize = myPageSize
	maxPages = cint(rs.pageCount)
	maxRecs = cint(rs.pageSize)
	rs.absolutePage = myPage
	howManyRecs = 0
	rec = 1
	do until rs.Eof or rec = 31 
		if i = 1 then 
			CColor = strAltForumCellColor
		else
			CColor = strForumCellColor
		end if
%>
      <tr>

        <td bgcolor="<% =CColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="pop_profile.asp?mode=display&id=<% =rs("MEMBER_ID") %>" target="_new"><b><% =ChkString(rs("M_NAME"),"display") %></a>&nbsp;&nbsp;<img src="<%=strImageURL %>pm.gif" width="11" height="17" alt="�������Ļ�ѶϢ" border="0" onClick="opener.document.PostTopic.sendto.value+='<% =ChkString(rs("M_NAME"),"display") %>'" style="cursor:hand"></font></td>

      </tr>
<%
		currMember = rs("MEMBER_ID")
		rs.MoveNext
		i = i + 1
		if i = 2 then i = 0
		rec = rec + 1
	loop 
end if 
%>

    </table>

<table cellpadding=4 cellspacing=0>
 <tr bgcolor="<% =strHeadCellColor %>">
 <form action="pm_pop_members.asp?mode=search" method="post" name="SearchMembers">
  <td><font face="����, Arial, Helvetica" size="2" color="#FFFFFF" size"+1"><b>������</b></font>&nbsp;<input type="text" name="M_NAME" size="14">&nbsp;
  </td>
  <td>
   <INPUT type="submit" value="����" value="search" id=submit1 name=submit1>
  </td>
 </tr> 
 </form> 
  <tr bgcolor="<% =strHeadCellColor %>">
    <td colspan="2" align="center" valign="top">        
	<a href="pm_pop_members.asp?mode=search&M_NAME=A&View=Calendar"><font color="#FFFFFF" face=arial size=1>A</font></a>
	<a href="pm_pop_members.asp?mode=search&M_NAME=B&View=Calendar"><font color="#FFFFFF"" face=arial size=1>B</font></a>
	<a href="pm_pop_members.asp?mode=search&M_NAME=C&View=Calendar"><font color="#FFFFFF" face=arial size=1>C</font></a>
	<a href="pm_pop_members.asp?mode=search&M_NAME=D&View=Calendar"><font color="#FFFFFF" face=arial size=1>D</font></a>
	<a href="pm_pop_members.asp?mode=search&M_NAME=E&View=Calendar"><font color="#FFFFFF" face=arial size=1>E</font></a>
	<a href="pm_pop_members.asp?mode=search&M_NAME=F&View=Calendar"><font color="#FFFFFF" face=arial size=1>F</font></a>
	<a href="pm_pop_members.asp?mode=search&M_NAME=G&View=Calendar"><font color="#FFFFFF" face=arial size=1>G</font></a>
	<a href="pm_pop_members.asp?mode=search&M_NAME=H&View=Calendar"><font color="#FFFFFF" face=arial size=1>H</font></a>
	<a href="pm_pop_members.asp?mode=search&M_NAME=I&View=Calendar"><font color="#FFFFFF" face=arial size=1>I</font></a>
	<a href="pm_pop_members.asp?mode=search&M_NAME=J&View=Calendar"><font color="#FFFFFF" face=arial size=1>J</font></a>
	<a href="pm_pop_members.asp?mode=search&M_NAME=K&View=Calendar"><font color="#FFFFFF" face=arial size=1>K</font></a>
	<a href="pm_pop_members.asp?mode=search&M_NAME=L&View=Calendar"><font color="#FFFFFF" face=arial size=1>L</font></a>
	<a href="pm_pop_members.asp?mode=search&M_NAME=M&View=Calendar"><font color="#FFFFFF" face=arial size=1>M</font></a>
	<a href="pm_pop_members.asp?mode=search&M_NAME=N&View=Calendar"><font color="#FFFFFF" face=arial size=1>N</font></a><br>
	<a href="pm_pop_members.asp?mode=search&M_NAME=O&View=Calendar"><font color="#FFFFFF" face=arial size=1>O</font></a>
	<a href="pm_pop_members.asp?mode=search&M_NAME=P&View=Calendar"><font color="#FFFFFF" face=arial size=1>P</font></a>
	<a href="pm_pop_members.asp?mode=search&M_NAME=Q&View=Calendar"><font color="#FFFFFF" face=arial size=1>Q</font></a>
	<a href="pm_pop_members.asp?mode=search&M_NAME=R&View=Calendar"><font color="#FFFFFF" face=arial size=1>R</font></a>
	<a href="pm_pop_members.asp?mode=search&M_NAME=S&View=Calendar"><font color="#FFFFFF" face=arial size=1>S</font></a>
	<a href="pm_pop_members.asp?mode=search&M_NAME=T&View=Calendar"><font color="#FFFFFF" face=arial size=1>T</font></a>
	<a href="pm_pop_members.asp?mode=search&M_NAME=U&View=Calendar"><font color="#FFFFFF" face=arial size=1>U</font></a>
	<a href="pm_pop_members.asp?mode=search&M_NAME=V&View=Calendar"><font color="#FFFFFF" face=arial size=1>V</font></a>
	<a href="pm_pop_members.asp?mode=search&M_NAME=W&View=Calendar"><font color="#FFFFFF" face=arial size=1>W</font></a>
	<a href="pm_pop_members.asp?mode=search&M_NAME=X&View=Calendar"><font color="#FFFFFF" face=arial size=1>X</font></a>
	<a href="pm_pop_members.asp?mode=search&M_NAME=Y&View=Calendar"><font color="#FFFFFF" face=arial size=1>Y</font></a>
	<a href="pm_pop_members.asp?mode=search&M_NAME=Z&View=Calendar"><font color="#FFFFFF" face=arial size=1>Z</font></a><br>
	</td>
  </tr>
</table>
    </td>
  </tr>
  <tr>
    <td>
    <table border="0" width="100%">
      <tr>
        <td>
<% if maxpages > 1 then %>
        <table border=0>
          <tr>
            <td valign="top"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>��Ա��</b></font></td>
            <td valign="top"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% Call Paging() %></font></td>
          </tr>
        </table>
<% else %>
        &nbsp;
<% end if %>
        </td>
      </tr>
    </table>
    </td>
  </tr>
</table>

<!--#INCLUDE FILE="inc_footer_short.asp" -->
<% end if %>
<%
sub Paging()
	if maxpages > 1 then
		if Request.QueryString("whichpage") = "" then
			sPageNumber = 1
		else
			sPageNumber = Request.QueryString("whichpage")
		end if
		if Request.QueryString("method") = "" then
			sMethod = "postsdesc"
		else
			sMethod = Request.QueryString("method")
		end if
		sScriptName = Request.ServerVariables("script_name")
		Response.Write("<table border=0 width=100% cellspacing=0 cellpadding=1 align=top><tr>")
		for counter = 1 to maxpages
			if counter <> cint(sPageNumber) then   
				sNum = "<td align=right bgcolor=" & strPageBGColor & "><font face=" & strDefaultFontFace & " size=" & strDefaultFontSize & ">" & "&nbsp;" & widenum(counter) &  "<a href=""" & sScriptName 
				sNum = sNum & "?whichpage=" & counter
				sNum = sNum & "&pagesize=" & mypagesize
				sNum = sNum & "&method=" & sMethod
				sNum = sNum & """>" & counter & "</a></font></td>"
				Response.Write sNum
			else
				Response.Write("<td align=right bgcolor=" & strPageBGColor & "><font face=" & strDefaultFontFace & " size=" & strDefaultFontSize & ">" & "&nbsp;" & widenum(counter) & "<b>" & counter & "</b></font></td>")
			end if
			if counter mod 30 = 0 then
				Response.Write("</tr><tr>")
			end if
		next
		Response.Write("</tr></table>")
	end if
end sub 
%>