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
<% 
	If Session(strCookieURL & "Approval") = "15916941253" Then %>
<!--#INCLUDE FILE="inc_functions.asp" -->
<!--#include file="inc_top.asp"-->
<%	

	if intUserFields = "" then intUserFields = 0
	if Request.Form("Method_Type") = "Write_Configuration" then 
		Err_Msg = ""
	
		if Err_Msg = "" then

			'## Forum_SQL
			if cInt(Request.Form("intUserFields")) <> 1 then
				intUserFields = 0
			else
				intUserFields = 1
			end if
			strSql = "UPDATE " & strTablePrefix & "MODS "
			strSql = strSql & " SET M_VALUE ='" & intUserFields & "'"
			strSQL = strSql & " WHERE M_NAME = 'HModEnable' AND M_CODE = 'UserFields' "

			my_Conn.Execute (strSql)

			Application(strCookieURL & "ConfigLoaded") = ""

%>
<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">�趨�Ѿ����������ϣ�</font></p>
<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="admin_user_fields.asp">�ص���Ա������λ������������</font></a></p>
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
	else 
%>

<table width="80%" height="80%" align="center">
<tr>
	<td align="center"><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>>
			<form action="admin_user_fields.asp" method="post" name="Addform">
			<input type="hidden" name="Method_Type" value="Write_Configuration">
			<% 
			if intUserFields <> 1 then %>
			<input type="checkbox" name="intUserFields" value="1" <% if intUserFields = 1 then response.write "checked" end if %> >
			�����Ա��������δ�������빴ѡ����Ŀ�������趨
			<% Else  %>
			<input type="checkbox" name="intUserFields" value="1" <% if intUserFields = 1 then response.write "checked" end if %> >&nbsp;
			�����Ա����������������ȡ����ѡ����Ŀ�������趨
			<% End If %>
			<br><input type="submit" name="AddField" value=" �趨 " style="font-weight: bold; font-size: x-small; color: White; background-color: #7b68ee;">
			</form></font>
	</td>
</tr>
  <tr>
    <td align=center valign=center>
    <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">

<%
			
			Set rs2 = Server.CreateObject("ADODB.Recordset")
			rs2.cachesize = 20
			sSQL = "SELECT * FROM " & strMemberTablePrefix & "USERFIELDS" 
			rs2.Open sSQL, my_Conn, 3
			rs2.cachesize = 5
			rs2.PageSize = 5
			strPageSize = 5
			lRecs = rs2.RecordCount
			lFields = rs2.Fields.count
			
			iPage = CLng(request("iPage"))

			iPageCount = cInt(rs2.PageCount)

			if iPage < 1 then
				iPage = 1
			end if
			if not rs2.BOF or not rs2.EOF then
				rs2.AbsolutePage = iPage
			end if

			%>
			<table align="center" border=0 cellspacing=1 cellpadding=4 bgcolor = "<% =strHeadCellColor %>" width=80%>
				<tr>
					
					<td bgcolor="<% =strHeadCellColor %>"  colspan=<%=lFields+1%>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">��λ&nbsp;[<%=lRecs%>]&nbsp;&nbsp;
						<%
							if iPage > 1 then
								response.write "<a href=""admin_user_fields.asp?uTable=FORUM_USERFIELDS&iPage=" & iPage - 1 & """ alt=""��һҳ"">"
								response.write "��</a>&nbsp;|&nbsp;"
							else
								response.write "��&nbsp;|&nbsp;"
							end if
							if iPage < iPageCount then 
								response.write "<a href=""admin_user_fields.asp?uTable=FORUM_USERFIELDS&iPage=" & iPage + 1 & """ alt=""��һҳ"">"
								response.write "��</a>"
							else
								response.write "��"
							end if
							response.write "&nbsp;&nbsp;&nbsp;[Page:" & iPage & "/" 
							if iPageCount = 0 then
								iPageCount = 1
							end if
							response.write iPageCount & "]"
					%>
					</td>
					
				</tr>
				<tr>
				<td bgcolor="<% =strPopUpTableColor %>" ></td>
				<td bgcolor="<% =strPopUpTableColor %>" ><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">��ʾ����</font></td>
				<td bgcolor="<% =strPopUpTableColor %>" ><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�����ͱ�</font></td>
				<td bgcolor="<% =strPopUpTableColor %>" ><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">��λ����</font></td>
				
			<%


			response.write "</tr>"
			
			rec = 1
			do until rs2.eof or rec = (strPageSize +1)
				
				set fld = rs2.fields(0)
				iFieldCount = 0
				response.write "<td bgcolor=" & strPopUpTableColor & "><font face=" & strDefaultFontFace & " size="& strDefaultFontSize & ">"
				response.write "<a href=""mods/pop_ufield_table.asp?uTable=FORUM_USERFIELDS&" & fld.name & "=" & rs2(fld.name) & "&iPage=" & iPage & "&preview=preDel"" >ɾ��</a>&nbsp;"
				response.write "<a href=""mods/pop_ufield_record.asp?uTable=FORUM_USERFIELDS&" & fld.name & "=" & rs2(fld.name) & "&iPage=" & iPage & """ >�༭</a>"
				response.write "</td>"			
				
				
				for each fld in rs2.fields
					iFieldCount = iFieldCount + 1
					if iFieldCount <> 1 then
					response.write "<td bgcolor=" & strPopUpTableColor & "><font face=" & strDefaultFontFace & " size="& strDefaultFontSize & ">"
							response.write rs2(fld.name)
					end if
					response.write "</font></td>"
				next
				
				response.write "</tr>"
				rec = rec + 1
				rs2.movenext
			loop
			response.write "</table>"
			rs2.close
			set rs2 = nothing


%>
			<form action="mods/pop_ufield_table.asp?preview=preNew" method="post" name="Addform">
			<input type="submit" name="AddField" value=" ���� " style="font-weight: bold; font-size: x-small; color: White; background-color: #7b68ee;">
			<input type="hidden" name="Method_Type" value="">
			</form>

<p><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>" >
<b>ɾ��</b> ��ѡɾ�������Ŀ<br>
<b>����</b> ���˰�ť������Ŀ<br>
</p>
    <p><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="admin_home.asp">������̳��������</A></font></p>
    </font>
    </center>
	</tr>
	</table>
			</div>
</body>
</html>
<% End If %>	
<!--#include file="inc_footer.asp"-->
<% Else %>
<% Response.Redirect "admin_login.asp" %>
<% End IF %>
