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
<%
'#################################################################################
strRqMethod = Request.QueryString("method")
'#################################################################################

if strRqMethod = "Topic" then

	if Request.QueryString("TOPIC_ID") = ""  then
		Response.Redirect "default.asp"
		Response.End
	end if
        
        set my_Conn = Server.CreateObject("ADODB.Connection") 
	my_Conn.Open strConnString
        	
	if (strAuthType = "nt") then
		set my_Conn = Server.CreateObject("ADODB.Connection")
		my_Conn.Open strConnString
		call NTauthenticate()
		if (ChkAccountReg() = "1") then
			call NTUser()
		end if
	end if

	'## Forum_SQL - Find out if the Topic is Locked or Un-Locked and if it Exists
	strSql = "SELECT " & strTablePrefix & "TOPICS.CAT_ID, " & strTablePrefix & "TOPICS.FORUM_ID, " & strTablePrefix & "TOPICS.TOPIC_ID, " & strTablePrefix & "TOPICS.T_SUBJECT, " & strTablePrefix & "FORUM.F_SUBJECT " 
	strSql = strSql & " FROM " & strTablePrefix & "TOPICS, " & strTablePrefix & "FORUM "
	strSql = strSql & " WHERE " & strTablePrefix & "TOPICS.TOPIC_ID = " & Request.QueryString("TOPIC_ID")
	strSql = strSql & " AND " & strTablePrefix & "TOPICS.FORUM_ID = " & strTablePrefix & "FORUM.FORUM_ID"

	set rsTopicInfo = my_Conn.Execute (StrSql)

	if (rsTopicInfo.EOF and rsTopicInfo.BOF) then
		Response.Redirect "default.asp"
	else
		Response.Redirect "topic.asp?TOPIC_ID=" & RsTopicInfo("TOPIC_ID") & "&FORUM_ID=" & RsTopicInfo("FORUM_ID") & "&CAT_ID=" & RsTopicInfo("CAT_ID") & "&Forum_Title=" & ChkString(RsTopicInfo("F_SUBJECT"),"urlpath") & "&Topic_Title=" & ChkString(RsTopicInfo("T_SUBJECT"),"urlpath")
	end if
end if

if strRqMethod = "Forum" then

	if Request.QueryString("FORUM_ID") = ""  then
		Response.Redirect "default.asp"
		Response.End
	end if

        set my_Conn = Server.CreateObject("ADODB.Connection") 
	my_Conn.Open strConnString
        	
	if (strAuthType = "nt") then
		set my_Conn = Server.CreateObject("ADODB.Connection")
		my_Conn.Open strConnString
		call NTauthenticate()
		if (ChkAccountReg() = "1") then
			call NTUser()
		end if
	end if

	'## Forum_SQL - Find out if the Topic is Locked or Un-Locked and if it Exists
	strSql = "SELECT " & strTablePrefix & "FORUM.FORUM_ID, " & strTablePrefix & "FORUM.CAT_ID, " & strTablePrefix & "FORUM.F_SUBJECT, " & strTablePrefix & "FORUM.CAT_ID, " & strTablePrefix & "FORUM.F_SUBJECT " 
	strSql = strSql & " FROM " & strTablePrefix & "FORUM "
	strSql = strSql & " WHERE " & strTablePrefix & "FORUM.FORUM_ID = " & Request.QueryString("FORUM_ID")

	set rsForumInfo = my_Conn.Execute (StrSql)

	if (rsForumInfo.EOF and rsForumInfo.BOF) then
		Response.Redirect "default.asp"
	else
	        Response.Redirect "forum.asp?FORUM_ID=" & rsForumInfo("FORUM_ID") & "&CAT_ID=" & rsForumInfo("CAT_ID") & "&Forum_Title=" & ChkString(rsForumInfo("F_SUBJECT"),"urlpath")
	end if
end if
%>