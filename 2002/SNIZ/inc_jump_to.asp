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
<!-- JUMP TO --> 
    <form name="Stuff">
    <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">������</font>
    <select name="SelectMenu" size="1" onchange="jumpTo(this)">
<!--    <select name="SelectMenu" size="1" onchange="jumpTo(document.Stuff.SelectMenu)">-->
      <option value="./">��ѡ��һ����̳</option>
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
      <option value="<% =strHomeURL %>">������ҳ
      <option value="active.asp">��������
      <option value="faq.asp">����˵��
      <option value="members.asp">��Ա�б�
      <option value="search.asp">��̳����
    </select>
    </form>
<!-- END JUMP TO -->
