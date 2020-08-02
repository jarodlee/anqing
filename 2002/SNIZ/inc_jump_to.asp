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
<!-- JUMP TO --> 
    <form name="Stuff">
    <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">跳至：</font>
    <select name="SelectMenu" size="1" onchange="jumpTo(this)">
<!--    <select name="SelectMenu" size="1" onchange="jumpTo(document.Stuff.SelectMenu)">-->
      <option value="./">请选择一个论坛</option>
<%
	'## Get all Forum Categories From DB
	strSql = "SELECT " & strTablePrefix & "CATEGORY.CAT_ID, " & strTablePrefix & "CATEGORY.CAT_NAME "
	strSql = strSql & " FROM " & strTablePrefix & "CATEGORY"
	strSql = strSql & " ORDER BY " & strTablePrefix & "CATEGORY.CAT_NAME"
	
	set rsCat = my_conn.Execute (strSql)

	do until rsCat.eof '## Grab the Categories.

		'##  Build SQL to get forums via category
		'## Forum_SQL
		strSql = "SELECT " & strTablePrefix & "FORUM.FORUM_ID, " & strTablePrefix & "FORUM.F_TYPE, " & strTablePrefix & "FORUM.F_SUBJECT, " & strTablePrefix & "FORUM.F_URL, " & strTablePrefix & "FORUM.CAT_ID "
		strSql = strSql & " FROM " & strTablePrefix & "FORUM "
		strSql = strSql & " WHERE " & strTablePrefix & "FORUM.CAT_ID = " & rsCat("CAT_ID")
		strSql = strSql & " ORDER BY " & strTablePrefix & "FORUM.F_SUBJECT ASC;"

		set rsForum =  my_Conn.Execute (StrSql)

		if rsForum.eof or rsForum.bof then
			'nothing
		else
			iNewCat = rsForum("CAT_ID")
			iOldCat = 0
			do until rsForum.Eof
				if ChkForumAccess(rsForum("FORUM_ID")) then
					if iNewCat <> iOldCat Then
						Response.Write "      <option value='default.asp'>" & rsCat("CAT_NAME") & "</option>" & vbCrLf
						iOldCat = iNewCat
					end if
					if rsForum("F_TYPE") = 0 then
						Response.Write "      <option value='forum.asp?FORUM_ID=" & rsForum("FORUM_ID") & "&CAT_ID=" & rsForum("CAT_ID") & "&Forum_Title=" & ChkString(rsForum("F_SUBJECT"),"urlpath") & "'"
					else
						if rsForum("F_TYPE") = 1 then
							Response.Write "      <option value='" & rsForum("F_URL") & "'"
						end if
					end if
					if rsForum("FORUM_ID") = Request.Querystring("Forum_ID") then 
						Response.Write(" SELECTED") 
					end if
					Response.Write ">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & rsForum("F_SUBJECT")& "</option>" & vbCrLf
				end if
				rsForum.MoveNext
			loop
		end if
		rsCat.MoveNext
	loop
%>
      <option value="">&nbsp;--------------------
      <option value="<% =strHomeURL %>">返回主页
      <option value="active.asp">最新文章
      <option value="faq.asp">帮助说明
      <option value="members.asp">会员列表
      <option value="search.asp">论坛搜索
    </select>
    </form>
<!-- END JUMP TO -->
