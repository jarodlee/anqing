<%
'########## Snitz Forums 2000 Version 3.1 SR4 ####################
'#                                                               #
'#  汉化修改: 资源搜罗站                                         #
'#  电子邮件: cgier@21cn.com                                     #
'#  主页地址: http://www.sdsea.com                               #
'#            http://www.99ss.net                                #
'#            http://www.cdown.net                               #
'#	     http://www.wzdown.com                               #
'#	     http://www.13888.net                                #
'#  论坛地址:http://ubb.yesky.net                                #
'#  最后修改日期: 2001/03/12    中文版本：Version 3.1 SR4        #
'#################################################################
'# 原始来源                                                      #
'# Snitz Forums 2000 Version 3.1 SR4                             #
'# Copyright 2000 http://forum.snitz.com - All Rights Reserved   #
'#################################################################
'#【版权声明】                                                   #
'#                                                               #
'# 本软体为共享软体(shareware)提供个人网站免费使用，请勿非法修改,#
'# 转载，散播，或用于其他图利行为，并请勿删除版权声明。          #
'# 如果您的网站正式起用了这个脚本，请您通知我们，以便我们能够知晓#
'# 如果可能，请在您的网站做上我们的链接，希望能给予合作。谢谢！  #
'#################################################################
'# 请您尊重我们的劳动和版权，不要删除以上的版权声明部分，谢谢合作#
'# 如有任何问题请到我们的论坛告诉我们                            #
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
<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">设定已经更新完成完毕！</font></p>
<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="admin_user_fields.asp">回到会员资料栏位新增管理中心</font></a></p>
<%		else %>
	<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">你输入的资料有问题或没有填写完整</font></p>

	<table align=center border=0>
	  <tr>
	    <td><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
		<ul>
		<% =Err_Msg %>
		</ul>
	    </font></td>
	  </tr>
	</table>

	<p align=center><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>><a href="JavaScript:history.go(-1)">请返回重新输入</a></font></p>
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
			额外会员资料栏并未启动，请勾选该项目并按下设定
			<% Else  %>
			<input type="checkbox" name="intUserFields" value="1" <% if intUserFields = 1 then response.write "checked" end if %> >&nbsp;
			额外会员资料栏已启动，请取消勾选该项目并按下设定
			<% End If %>
			<br><input type="submit" name="AddField" value=" 设定 " style="font-weight: bold; font-size: x-small; color: White; background-color: #7b68ee;">
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
					
					<td bgcolor="<% =strHeadCellColor %>"  colspan=<%=lFields+1%>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">栏位&nbsp;[<%=lRecs%>]&nbsp;&nbsp;
						<%
							if iPage > 1 then
								response.write "<a href=""admin_user_fields.asp?uTable=FORUM_USERFIELDS&iPage=" & iPage - 1 & """ alt=""上一页"">"
								response.write "←</a>&nbsp;|&nbsp;"
							else
								response.write "←&nbsp;|&nbsp;"
							end if
							if iPage < iPageCount then 
								response.write "<a href=""admin_user_fields.asp?uTable=FORUM_USERFIELDS&iPage=" & iPage + 1 & """ alt=""下一页"">"
								response.write "→</a>"
							else
								response.write "→"
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
				<td bgcolor="<% =strPopUpTableColor %>" ><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">显示名称</font></td>
				<td bgcolor="<% =strPopUpTableColor %>" ><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">资料型别</font></td>
				<td bgcolor="<% =strPopUpTableColor %>" ><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">栏位名称</font></td>
				
			<%


			response.write "</tr>"
			
			rec = 1
			do until rs2.eof or rec = (strPageSize +1)
				
				set fld = rs2.fields(0)
				iFieldCount = 0
				response.write "<td bgcolor=" & strPopUpTableColor & "><font face=" & strDefaultFontFace & " size="& strDefaultFontSize & ">"
				response.write "<a href=""mods/pop_ufield_table.asp?uTable=FORUM_USERFIELDS&" & fld.name & "=" & rs2(fld.name) & "&iPage=" & iPage & "&preview=preDel"" >删除</a>&nbsp;"
				response.write "<a href=""mods/pop_ufield_record.asp?uTable=FORUM_USERFIELDS&" & fld.name & "=" & rs2(fld.name) & "&iPage=" & iPage & """ >编辑</a>"
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
			<input type="submit" name="AddField" value=" 新增 " style="font-weight: bold; font-size: x-small; color: White; background-color: #7b68ee;">
			<input type="hidden" name="Method_Type" value="">
			</form>

<p><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>" >
<b>删除</b> 点选删除这个栏目<br>
<b>新增</b> 按此按钮新增栏目<br>
</p>
    <p><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="admin_home.asp">返回论坛管理中心</A></font></p>
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
