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
<% If Session(strCookieURL & "Approval") = "15916941253" Then %>
<!--#INCLUDE file="inc_functions.asp" -->
<!--#INCLUDE file="inc_top.asp" -->
<table border="0" width="100%">
  <tr>
    <td width="33%" align="left" nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    <img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="default.asp">������̳��ҳ</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="admin_home.asp">��̳��������</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" height=15 width=15 border="0">&nbsp;�ʼ������趨<br>
    </font></td>
  </tr>
</table>
<%
	if Request.Form("Method_Type") = "Write_Configuration" then 
		Err_Msg = ""
		if Request.Form("strMailServer") = "" and Request.Form("strEmail") = "1" then 
			Err_Msg = Err_Msg & "<li>����������ʼ�����λַ��</li>"
		end if
		if ((lcase(left(Request.Form("strMailServer"), 7)) = "http://") or (lcase(left(Request.Form("strMailServer"), 8)) = "https://") or Request.Form("strMailServer") = "") and Request.Form("strEmail") = "1" then
			Err_Msg = Err_Msg & "<li>��Ҫ���ʼ�����λַǰ���� <b>http://</b>, <b>https://</b> or <b>file://</b></li>"
		end if
		if Request.Form("strSender") = "" then 
			Err_Msg = Err_Msg & "<li>������������Ա�ĵ����ʼ�λַ</li>"
		else
			if EmailField(Request.Form("strSender")) = 0 and Request.Form("strSender") <> "" then 
				Err_Msg = Err_Msg & "<li>���������Ϸ��Ĺ���Ա�����ʼ���ַ</li>"
			end if
		end if
	
		if Err_Msg = "" then

			'## Forum_SQL
			strSql = "UPDATE " & strTablePrefix & "CONFIG "
			strSql = strSql & " SET C_STREMAIL                 = " & Request.Form("strEmail") & ""
			strSql = strSql & ",    C_STRMAILMODE              = '" & Request.Form("strMailMode") & "'"
			if Request.Form("strMailServer") <> "" then
				strSql = strSql & ",    C_STRMAILSERVER            = '" & Request.Form("strMailServer") & "'"
			end if
			if Request.Form("strSender") <> "" then
				strSql = strSql & ",    C_STRSENDER                = '" & Request.Form("strSender") & "'"
			end if
			strSql = strSql & ",    C_STRUNIQUEEMAIL           = " & Request.Form("strUniqueEmail") & ""
			strSql = strSql & ",    C_STRLOGONFORMAIL          = " & Request.Form("strLogonForMail") & ""
		
			my_Conn.Execute(strSql)
			Application(strCookieURL & "ConfigLoaded") = ""

%>
<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">�ʼ������趨�Ѿ�������ϣ�</font></p>
<meta http-equiv="Refresh" content="2; URL=admin_home.asp">

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">�޸���ɣ�</font></p>

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="admin_home.asp">�ص���������</font></a></p>
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
<%		end if %>
<%	else %>
<form action="admin_config_email.asp" method="post" id="Form1" name="Form1">
<input type="hidden" name="Method_Type" value="Write_Configuration">
<table border="0" cellspacing="1" cellpadding="4" align=center bgcolor="<% =strPopUpBorderColor %>">
  <tr valign="top">
    <td bgcolor="<% =strHeadCellColor %>" colspan="2"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>"><b>�ʼ������趨</b></font></td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�ʼ�������ģʽ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    ������<input type="radio" class="radio" name="strEmail" value="1" <% if lcase(strEmail) <> "0" then Response.Write("checked") %>> 
    �رգ�<input type="radio" class="radio" name="strEmail" value="0" <% if lcase(strEmail) = "0" then Response.Write("checked") %>>
    <select name="strMailMode">
      <option value="aspmail"<% if (lcase(strMailMode)="aspmail" or lcase(MailMode)="aspmail") then Response.Write(" selected") %>>ASPMail</option>
      <option value="aspemail"<% if (lcase(strMailMode)="aspemail" or lcase(MailMode)="aspemail") then Response.Write(" selected") %>>ASPEMail</option>
      <option value="aspqmail"<% if (lcase(strMailMode)="aspqmail" or lcase(MailMode)="aspqmail") then Response.Write(" selected") %>>ASPQMail</option>
      <option value="cdonts"<% if (lcase(strMailMode)="cdonts" or lcase(MailMode)="cdonts") then Response.Write(" selected") %>>CDONTS</option>
      <option value="chilicdonts"<% if (lcase(strMailMode)="chilicdonts" or lcase(MailMode)="chilicdonts") then Response.Write(" selected") %>>Chili!Mail</option>
      <option value="dkqmail"<% if (lcase(strMailMode)="dkqmail") then Response.Write(" selected") %>>dkQMail</option>
      <option value="geocel"<% if (lcase(strMailMode)="geocel" or lcase(MailMode)="geocel") then Response.Write(" selected") %>>GeoCel</option>
      <option value="iismail"<% if (lcase(strMailMode)="iismail" or lcase(MailMode)="iismail") then Response.Write(" selected") %>>IISMail</option>
      <option value="jmail"<% if (lcase(strMailMode)="jmail" or lcase(MailMode)="jmail") then Response.Write(" selected") %>>JMail</option>
      <option value="smtp"<% if (lcase(strMailMode)="smtp" or lcase(MailMode)="smtp") then Response.Write(" selected") %>>SMTP</option>
    </select>
    <a href="JavaScript:openWindow3('pop_config_help.asp#email')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </font></td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�ʼ���������ַ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    <input type="text" name="strMailServer" size="25" value="<% =strMailServer %>">
    <a href="JavaScript:openWindow3('pop_config_help.asp#mailserver')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">��̳����Ա�ʼ���ַ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    <input type="text" name="strSender" size="25" value="<% =strSender %>">
    <a href="JavaScript:openWindow3('pop_config_help.asp#sender')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�ʼ���ַ�����ظ�ע�᣺</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    ������<input type="radio" class="radio" name="strUniqueEmail" value="1" <% if strUniqueEmail = "1" then Response.Write("checked") %>> 
    �رգ�<input type="radio" class="radio" name="strUniqueEmail" value="0" <% if strUniqueEmail <> "1" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#UniqueEmail')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </font></td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">���͵����ʼ���</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    ������<input type="radio" class="radio" name="strLogonForMail" value="1" <% if strLogonForMail = "1" then Response.Write("checked") %>> 
    �رգ�<input type="radio" class="radio" name="strLogonForMail" value="0" <% if strLogonForMail <> "1" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#LogonForMail')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </font></td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" colspan="2" align="center"><input type="submit" value="�ύ����" id="submit1" name="submit1"> <input type="reset" value="��������" id="reset1" name="reset1"></td>
  </tr>
</table>
</form>
</font>
<%	end if %>
<!--#INCLUDE file="inc_footer.asp" -->
<% Else %>
<% Response.Redirect "admin_login.asp" %>
<% End IF %>
