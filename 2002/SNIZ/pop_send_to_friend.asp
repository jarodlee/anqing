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
<script language="JavaScript">
 function changeTopics(){
 
	document.Form2.Tcount.value = document.Form1.TopicCount.value
	document.Form2.submit()
}
</script> 

<%  if Request.QueryString("mode") <> "DoIt" then
	if Request.QueryString("TOPIC_ID") = "" then
		TopicID = Request("hTopicID")
	else
		TopicID = Request.QueryString("TOPIC_ID")
	end if
	'## Forum_SQL - Get Origional Posting
	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strTablePrefix & "TOPICS.T_DATE, " & strTablePrefix & "TOPICS.T_SUBJECT, " & strTablePrefix & "TOPICS.T_AUTHOR, " & strTablePrefix & "TOPICS.TOPIC_ID, " & strTablePrefix & "TOPICS.T_MESSAGE "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS, " & strTablePrefix & "TOPICS "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "TOPICS.T_AUTHOR "
	strSql = strSql & " AND   " & strTablePrefix & "TOPICS.TOPIC_ID = " &  TopicID 
	set rs4 = my_Conn.Execute (strSql)
	'## Forum_SQL - Get all topicsFrom DB
	strSql ="SELECT " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strTablePrefix & "REPLY.REPLY_ID, " & strTablePrefix & "REPLY.R_AUTHOR, " & strTablePrefix & "REPLY.TOPIC_ID, " & strTablePrefix & "REPLY.R_MESSAGE, " & strTablePrefix & "REPLY.R_DATE "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS, " & strTablePrefix & "REPLY "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "REPLY.R_AUTHOR "
	strSql = strSql & " AND   TOPIC_ID = " & TopicID & " "
	strSql = strSql & " ORDER BY " & strTablePrefix & "REPLY.R_DATE"
	set rs3 = Server.CreateObject("ADODB.Recordset")
	rs3.cachesize = 20
	rs3.open  strSql, my_Conn, 3
end if
		
if Request("Tcount") = ""	Then 
	TopicCount = 9999
else
	TopicCount = Request("Tcount")
end if
if Request.QueryString("mode") = "DoIt" then
	Err_Msg = ""
	if (Request.Form("YName") = "") then 
		Err_Msg = Err_Msg & "<li>������������������</li>"
	end if
	if (Request.Form("YEmail") = "") then 
		Err_Msg = Err_Msg & "<li>�������������ʼ���ַ��</li>"
	else
		if (EmailField(Request.Form("YEmail")) = 0) then 
			Err_Msg = Err_Msg & "<li>������д��ȷ���ʼ���ַ</li>"
		end if
	end if
	if (Request.Form("Name") = "") then 
		Err_Msg = Err_Msg & "<li>����������ռ���������</li>"
	end if
	if (Request.Form("Email") = "") then 
		Err_Msg = Err_Msg & "<li>����������ռ����ʼ���ַ��</li>"
	else
		if (EmailField(Request.Form("Email")) = 0) then 
			Err_Msg = Err_Msg & "<li>�����������ȷ���ռ����ʼ���ַ��</li>"
		end if
	end if
	if (Request.Form("Msg") = "") then 
		Err_Msg = Err_Msg & "<li>����������ʼ����ݣ�</li>"
	end if
	'##  Emails Topic To a Friend.  
	'##  This needs to be Edited to use your own email component
	'##  if you don't have one, try the w3Jmail component from www.dimac.net it's free!
	if lcase(strEmail) = "1" then
		if (Err_Msg = "") then
			strRecipientsName = Request.Form("Name")
			strRecipients = Request.Form("Email")
			strSubject = "From: " & Request.Form("YName") & " - Interesting Discussion at " & strForumTitle
			strMessage = "Hello " & Request.Form("Name") & vbCrLf & vbCrLf
			strMessage = strMessage & Request.Form("Msg") & vbCrLf & vbCrLf
			strMessage = strMessage & "You received this e-mail from: " & Request.Form("YName") & " " & Request.Form("YEmail")
%>
<!--#INCLUDE FILE="inc_mail.asp" -->
<%
			Response.Write "<p><font face='" & strDefaultFontFace & "' size=" & strHeaderFontSize & ">�ʼ��Ѽĳ���</font></p>" & vbCrLf
		else
%>
	<p><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">�������ʼ������⣬��������������û����д������</font></p>

	<table>
	  <tr>
	    <td><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
		<ul>
		<% =Err_Msg %>
		</ul></font>
	    </td>
	  </tr>
	</table>

	<p><font size="<% =strDefaultFontSize %>"><a href="JavaScript:history.go(-1)">�뷵����������</a></font></p>

	<p><font size="<% =strDefaultFontSize %>"><a href="JavaScript:onClick= window.close()">�رմ���</A></font></p>
<%
			Response.End
		end if
	end if
else 
%>


<% 


		'## Forum_SQL
		strSql =  "SELECT M_NAME, M_USERNAME, M_EMAIL "
		strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS"
		strSql = strSql & " WHERE "&Strdbntsqlname&" = '" & STRdbntUserName & "'"

		set rs = my_conn.Execute (strSql)
		YName = ""
		YEmail = ""

		if (rs.EOF or rs.BOF)  then
			if strLogonForMail <> "0" then 
			
				Err_Msg = Err_Msg & "<li>���͵����ʼ�ǰ�����ȵ�½��̳��</li>"
%>
			<p><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">�������ʼ������⣬��������������û����д������</font></p>

			<table>
			 <tr>
			  <td><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
			   <ul>
				<% =Err_Msg %>
			   </ul></font>
			 </td>
			 </tr>
			</table>

			<p><font size="<% =strDefaultFontSize %>"><a href="JavaScript:onClick= window.close()">�رմ���</A></font></p>
<%
				Response.End
			end if
		else
			YName = Trim("" & rs("M_NAME"))
			YEmail = Trim("" & rs("M_EMAIL"))
		end if
		rs.close
		set rs = nothing

%>

<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">���������Ƽ���������</font>
<form action="pop_send_to_friend.asp?<% =Request.QueryString %>" name=Form2 id=Form2><input type=hidden name="Tcount" value="">
<input type=hidden name="hTopicID" value="<%= TopicID %>">
</form>
<form action="pop_send_to_friend.asp?mode=DoIt" method=post id=Form1 name=Form1>
<input type=hidden name="Page" value="<% =Request.QueryString %>">
	
<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="<% =strPopUpBorderColor %>">
    <table border="0" width="100%" cellspacing="1" cellpadding="2">
      <TR>
        <TD bgColor=<% =strPopUpTableColor %> align="right" nowrap><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>>�ռ���������</font></td>
        <TD bgColor=<% =strPopUpTableColor %>><input type=text name="Name" size=25></td>
      </tr>
      <TR>
        <TD bgColor=<% =strPopUpTableColor %> align="right" nowrap><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>>�ռ����ʼ���ַ��</font></td>
        <TD bgColor=<% =strPopUpTableColor %>><input type=text name="Email" size=25></td>
      </tr>                
      <tr>
        <td bgColor=<% =strPopUpTableColor %> align="right" nowrap><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>>������������</font></td>
        <td bgColor=<% =strPopUpTableColor %>><input name=YName type=<% if YName <> "" then Response.Write("hidden") else Response.Write("text") end if %> value="<% = YName %>" size=25><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>> <% if YName <> "" then Response.Write(YName) end if %></font></td>
      </tr>
      <tr>
        <td bgColor=<% =strPopUpTableColor %> align="right" nowrap><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>>�������ʼ���ַ��</font></td>
        <td bgColor=<% =strPopUpTableColor %>><input name=YEmail type=<% if YEmail <> "" then Response.Write("hidden") else Response.Write("text") end if %> value="<% = YEmail %>" size=25><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>> <% if YEmail <> "" then Response.Write(YEmail) end if %></font></td>
      </tr> 

      <tr>
        <td bgColor=<% =strPopUpTableColor %> align="right" nowrap><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>>�ż����ݣ�</font></td>
		<td bgColor=<% =strPopUpTableColor %> nowrap><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>>Ҫ������ƪ�ظ���</font><input type="text" name="TopicCount" value="5" size="4"> <input type="button" value="���" name="change" onclick="changeTopics()"></td>
      </tr>
      <tr>
        <td bgColor=<% =strPopUpTableColor %> colspan=2 align=center><textarea name="Msg" cols="60" rows=10 readonly>�� <% =strForumTitle %> (<% =strForumURL%>) ����ҵ�һƪ���������Ȥ�����£�<% =vbCrLf & strForumURL%>/link.asp?TOPIC_ID=<%= TopicID & vbCrLf & vbCrLf %>������ƪ�����е�һ���ݣ�<% =vbCrLf %>
==========================================================

���⣺<% =rs4("T_Subject")%>
���ߣ�<% =rs4("M_NAME")%>
ʱ�䣺<% =ChkDate(rs4("T_DATE")) %>
���ݣ�<% =vbCrLf & rs4("T_MESSAGE") %>
<% i = 0
	rec=0 
	If rs3.EOF or rs3.BOF then  '## No categories found in DB
		Response.Write ""
	Else
		rs3.movefirst
		do until (rs3.EOF) or (rec > cInt(TopicCount))%>
----------------------------------------------------------
���ߣ�<% =rs3("M_NAME")%>
ʱ�䣺<% =ChkDate(rs3("R_DATE")) %>
���ݣ�<% =vbCrLf & rs3("R_MESSAGE") %>
<% 
rs3.MoveNext
		    strI = strI + 1
		    if strI = 2 then 
				strI = 0
			end if
		    rec = rec + 1
		loop
	end if
 %>
==========================================================
��ѶϢ���˽�����
<br>
��ӭ�������ǵ���̳��<% =strForumTitle%>
<% =strForumURL%>
</textarea></td>
      </tr>                    
      <tr>
        <td bgColor=<% =strPopUpTableColor %> colspan=2 align=center><input type=submit value="�ύ" id=Submit1 name=Submit1></td>
      </tr>
    </table>
    </td>
  </tr>
</table>
</form>
<% end if %>
<!--#INCLUDE FILE="inc_footer_short.asp" -->
