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
<!--#include file="config.asp"-->
<% 	If Session(strCookieURL & "Approval") = "15916941253" Then %>
<!--#include file="inc_pollmentor.asp"-->
<!--#include file="inc_functions.asp"-->
<!--#include file="inc_top.asp"-->
<%

	if intPollMentor = "" then intPollMentor = 0
	if Request.Form("Method_Type") = "Write_Configuration" then 
		Err_Msg = ""
	
		if Err_Msg = "" then

			'## Forum_SQL
			if cInt(Request.Form("intPollMentor")) <> 1 then
				intPollMentor = 0
			else
				intPollMentor = 1
			end if
			strSql = "UPDATE " & strTablePrefix & "MODS "
			strSql = strSql & " SET M_VALUE ='" & intPollMentor & "'"
			strSQL = strSql & " WHERE M_NAME = 'HModEnable' AND M_CODE = 'PollMentor' "

			my_Conn.Execute (strSql)

			Application(strCookieURL & "ConfigLoaded") = ""

%>
<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">�޸���ɣ�</font></p>
<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">�趨�ѱ��</font></p>
<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="admin_pollmentor.asp">����ͶƱ��������</font></a></p>
<%		else %>
	<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">������������������û����д����</font></p>

	<table align=center border=0>
	  <tr>
	    <td><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
		<ul>
		<% =Err_Msg %>
		</ul>
	    </font></td>
	  </tr>
	</table>

	<p align=center><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>><a href="JavaScript:history.go(-1)">�뷵����������</a></font></p>
<%		end if 
	else %>
			<div align="center"><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>>
			<form action="admin_pollmentor.asp" method="post" name="Addform">
			<input type="hidden" name="Method_Type" value="Write_Configuration">
			<% 
			if intPollMentor <> 1 then %>
			<input type="checkbox" name="intPollMentor" value="1" <% if intPollMentor = 1 then response.write "checked" end if %> >
			ͶƱ��Ŀ�����ǹرյģ���Ҫ����ͶƱ������ߴ򹴺������[����]
			<% Else  %>
			<input type="checkbox" name="intPollMentor" value="1" <% if intPollMentor = 1 then response.write "checked" end if %> >&nbsp;
			ͶƱ��Ŀ�����ǿ��ŵģ���Ҫ�ر�ͶƱ��ȥ����ߵĹ��������[����]
			<% End If %>
			<br><input type="submit" name="AddField" value="  �� ��  " style="font-weight: bold; font-size: x-small; color: White; background-color: #7b68ee;">
			</form></font></div>

<%'''''''''''  (C) Stefan Holmberg 1999  
'''''''''''  Free to use if these sourcecode lines is not deleted 
'''''''''''  Contact me at webmaster@sqlexperts.com
'''''''''''  http://www.sqlexperts.com
'''''''''''  AdMentor homepage at http://www.create-a-webshop.com

    Set oConn = my_conn
    
	Dim strTrue
	Dim strFalse
	
	If PollMentor_GetDatabaseType = "SQLServer" Then
		strTrue = 	"1"
		strFalse = "0"
	Else
		strTrue = "True"
		strFalse = "False"
	End If
	
	Dim strSort
	strSort = Request.QueryString("sort")    
	If strSort = "" Then
		strSort = " active, createdwhen desc"
	Else
		strSort = strSort & " desc"
	End If
	
	If Request.QueryString("saveactive") = "yes" Then
		Dim sActive
		Response.Write Request.Form("ACTIVE")
		oConn.Execute "update " & strTablePrefix  & "poll set active=" & strFalse & " where active=" & strTrue
		oConn.Execute "update " & strTablePrefix  & "poll set active=" & strTrue & " where id=" & Request.Form("ACTIVE")
	End If
	
%>


<br>
<table align="center" border="0" cellPadding="3" cellSpacing="1"  width="90%" bgcolor="<% =strPageBGColor %>">
  <tbody>
    <tr>
      <td vAlign="top" colspan="2" height="40">
        <table align="center" border="0" cellPadding="0" cellSpacing="1" bgcolor="<% =strTableBorderColor %>">
          <tbody>
            <tr>
              <td height="100%" vAlign="top" width="85%">
                <table bgcolor="<% =strForumcellColor %>" border="0" cellPadding="3" cellSpacing="1" height="100%" width="100%">
                  <tbody>
                    <tr>
                      <td align="left" height="100%" vAlign="top" width="65%">
                        <table border="0" width="100%">
                          <tr>
                            <td width="50%"><b><font color="#aa3333" face="verdana,arial,helvetica" size="4">ͶƱ�������� <%=Session("fullname")%></font></b>
                             
                            </td>
                            <td width="50%">
                            <%
                            Response.Write FAQ_GetAd(1)
                            %></td>
                          </tr>
                        </table>
                        <font color="#aa3333" face="verdana,arial,helvetica" size="+2">
                        <hr color="#000066" noShade SIZE="1">
                             
                        </font>
                             
                        <table border="0" width="100%">
                          <tr>
                            <td width="70%">
                             
                        <font face="helvetica, arial" size="2"><b>��ʱû��ͶƱ��Ŀ&nbsp;</b></font>
                              <form method="POST" action="admin_pollmentor.asp?saveactive=yes">
                        <table border="0" width="100%" cellspacing="1" cellpadding=3 bgcolor="<% =strTableBorderColor %>">
                          <tr bgcolor="<% =strCategoryCellColor %>">
                            <td>ѡ��</td>
                            <td><a href="admin_pollmentor.asp?sort=question">ͶƱ����</a></td>
                            <td><a href="admin_pollmentor.asp?sort=createdwhen">��ʼʱ��</a></td>
                            <td width="100">����</td>
                          </tr>
<%
	Dim oRS
	Set oRS = oConn.Execute("select * from " & strTablePrefix & "poll order by " & strSort)
	Dim bgcolor
	bgcolor = "#ECECD9"
	while not oRS.EOF
%>                        
                          <tr bgcolor="<% =strForumCellColor %>">
                          <%
                          Dim sBold, sBoldStop, sSel
                          If oRS("active") = True Then
                          	sBold = "<b>"
                          	sBoldStop ="</b>"
                          	sSel = "checked"
                          Else
                             sBold = ""
                             sBoldStop = ""
                             sSel = ""
                          End If
                          %>
                            <td bgcolor="<%=bgcolor%>">
                                <p><input type="radio" value="<%=oRS("id")%>" <%=sSel%> name="ACTIVE">
                              <%=sBold%></td>
                            <td bgcolor="<%=bgcolor%>"><%=sBold%><%=Trim(oRS("question"))%><%=sBoldStop%></td>
                            <td bgcolor="<%=bgcolor%>"><%=sBold%><%=oRS("createdwhen")%><%=sBoldStop%></td>
                            <td width="100" bgcolor="<%=bgcolor%>"><a href="admin_poll.asp?id=<%=oRS("id")%>&action=edit">�޸�</a>
                              - <a href="admin_poll.asp?id=<%=oRS("id")%>&save=yes&action=del">ɾ��</a></td>
                          </tr>
<%	
	If bgcolor= strForumCellColor  Then
		bgcolor = strAltForumCellColor
	Else
		bgcolor=strForumCellColor
	End if
	oRS.MoveNext
	Wend
	oRS.Close
	Set oRS = Nothing
%>                        
                        </table>
						<br>
                                <input type="submit" value="����ͶƱ��Ŀ" name="B1">
								<input type="hidden" name="Method_Type" value="">
                              </form>
                        <p><a href="admin_poll.asp?action=new">����һ���µ�ͶƱ��Ŀ</a></p>
                        <table border="0" width="100%">
                          <tr>
                            <td width="50%">
                             <%=FAQ_GetAd(3)
                             %>
                            </td>
                            <td width="50%">
                            <%=FAQ_GetAd(2)
                            %>
                            </td>
                          </tr>
                        </table>
                          </tr>
                        </table>
                      </td>
                    </tr>
                  </tbody>
                </table>
              </td>
            </tr>
          </tbody>
        </table>
      </td>
    </tr>
  </tbody>
</table>
<div align="center"><a href="admin_home.asp"><font face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>" size="<% =strDefaultFontSize %>">������̳��������</font></a></div>
<!--#include file="inc_footer.asp"-->
<% 	End if
	Else 
 	Response.Redirect "admin_login.asp"
 	End IF %>
