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
if Request.QueryString("mode") = "DoIt" then
	Err_Msg = ""
	if (Request.Form("Name") = "") then 
		Err_Msg = Err_Msg & "<li>�������������û���</li>"
	end if
	if (Request.Form("Email") = "") then 
		Err_Msg = Err_Msg & "<li>�����������ĵ����ʼ���ַ��</li>"
	end if
	if (EmailField(Request.Form("Email")) = 0) then 
		Err_Msg = Err_Msg & "<li>������д��ȷ�ĵ����ʼ���ַ</li>"
	end if
	'##  Emails Password to Requestor if Authenticated Properly.  
	'##  This needs to be Edited to use your own email component
	'##  if you don't have one, try the w3Jmail component from www.dimac.net it's free!
	if lcase(strEmail) = "1" then
		if (Err_Msg = "") then
			'## Forum_SQL
			strSql = "SELECT M_NAME, M_PASSWORD, M_EMAIL "
			strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
			strSql = strSql & " WHERE M_NAME = '" & Request.Form("Name") & "' AND "
			strSql = strSql & "       M_EMAIL = '" & Request.Form("email") & "'"

			set rs = my_Conn.Execute (strSql)

			strRecipientsName = Request.Form("Name")
			strRecipients = request.form("email")
			strSubject = strForumTitle & " " & "����Ѱ��"
			strMessage = "�������� " & strForumTitle & " ��ע�����붪ʧ����������ͨ��email���·���������㣬�����Ʊ��ܺ�..���ǵ���վ��ַ�� " & strForumURL & vbCrLf & vbCrLf

			if rs.EOF or rs.BOF then
				strMessage = strMessage & "Sorry. The details you supplied were wrong or unknown." & vbCrLf & vbCrLf
				strMessage = strMessage & "Please register again at " & strForumURL & "/register.asp?mode=Register" & vbCrLf
			Else
				strMessage = strMessage & " ���������: " & rs("M_PASSWORD") & vbCrLf
			end if
			
			strMessage = strMessage & "��ӭ���� " & strForumTitle
%>
<!--#INCLUDE FILE="inc_mail.asp" -->
<%
			rs.close
			set rs = nothing
'			Response.Write vbCrLf & "<br>������: " & strSender & vbCrLf & "<br>Recipients: " & strRecipients & vbCrLf & "<br>����: " & strSubject & vbCrLf & "<br>Message: " & strMessage & vbCrLf & "<br>"
			Response.Write "<p><font face=" & strDefaultFontFace & " size=" & strHeaderFontSize & ">�����Ѿ����ͻ�������䣡</font></p>" & vbCrLf
			Response.Write "<p><font face=" & strDefaultFontFace & " size=" & strDefaultFontSize & ">������ĵ�������.</font></p>" & vbCrLf
		else
%>
	<p><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">�������ʼ������⣬��������������û����д������</font></p>

	<table>
	  <tr>
	    <td>
		<ul>
		<% =Err_Msg %>
		</ul>
	    </td>
	  </tr>
	</table>

	<p><font size="<% =strDefaultFontSize %>"><a href="JavaScript:history.go(-1)">�뷵����������</a></font></p>

	<p><font size="<% =strDefaultFontSize %>"><a href="JavaScript:onClick= window.close()">�رմ���</A></font></p>
<%
			Response.End 
		end if
	end if
Else 
%>
<p><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">����Ѱ��</font></p>

<form action="pop_pword.asp?mode=DoIt" method="post">
<table border="0" width="75%" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="<% =strPopUpBorderColor %>">
    <table border="0" width="100%" cellspacing="1" cellpadding="1">
    <TBODY>
      <TR>
        <TD bgColor=<% =strPopUpTableColor %> align=right nowrap><b><FONT face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>>�û����ƣ�</font></b></TD>
        <TD bgColor=<% =strPopUpTableColor %>><INPUT name=Name size=25></TD>
      <TR>
        <TD bgColor=<% =strPopUpTableColor %> align=right nowrap><b><FONT face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>>�ʼ���ַ:</font></b></TD>
        <TD bgColor=<% =strPopUpTableColor %>><INPUT name=email size=25 value=""></TD>
      </TR>
        <TD bgColor=<% =strPopUpTableColor %> align=middle colSpan=2><INPUT name=submit1 type=submit value=�ύ></TD>
      </TR>
    </TBODY>
    </TABLE>
    </td>
  </tr>
</table>
</form>
<% end if %>
<!--#INCLUDE FILE="inc_footer_short.asp" -->
