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
    <img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="default.asp">������̳��ҳ</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="admin_home.asp">��̳��������</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" height=15 width=15 border="0">&nbsp;��̳��������<br>
    </font></td>
  </tr>
</table>
<%
	if Request.Form("Method_Type") = "Write_Configuration" then 
		Err_Msg = ""
		if Request.Form("strTitleImage") = "" then 
			Err_Msg = Err_Msg & "<li>�����������̳logo��ַ</li>"
		end if
		if Request.Form("strHomeURL") = "" then 
			Err_Msg = Err_Msg & "<li>�����������վ��ҳ�����·�����������ᣩ</li>"
		end if
		if Request.Form("strForumURL") = "" then 
			Err_Msg = Err_Msg & "<li>�����������̳��ַ</li>"
		end if
		if (left(lcase(Request.Form("strForumURL")), 7) <> "http://" and left(lcase(Request.Form("strForumURL")), 8) <> "https://") and Request.Form("strHomeURL") <> "" then
			Err_Msg = Err_Msg & "<li>�����������ǰ���� <b>http://</b>, <b>https://</b> or <b>file://</b></li>"
		end if
		if (right(lcase(Request.Form("strForumURL")), 1) <> "/") then
			Err_Msg = Err_Msg & "<li>���ӵ�ַ������ <b>/</b> Ϊ��β</li>"
		end if
		if Request.Form("strAuthType") <> strAuthType and strAuthType = "db" then 
			mLev = cint(ChkUser2(Request.Cookies(strUniqueID & "User")("Name"), Request.Cookies(strUniqueID & "User")("Pword")))
						
			if not(mLev = 4 and getMemberNumber(Request.Cookies(strUniqueID & "User")("Name")) = 1) then
				Err_Msg = Err_Msg & "<li>ֻ�й���Ա�����޸���̳����Ȩģʽ</li>"
			else
				call NTauthenticate()
				if Session(strCookieURL & "userid") = "" then
					Err_Msg = Err_Msg & "<li>ʹ����̳ǰ������������������ķ�������ȡģʽ</li>"
				else	
					strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
					strSql = strSql & " SET " & strMemberTablePrefix & "MEMBERS.M_USERNAME = '" & Session(strCookieURL & "userid") & "'"
					strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & Request.Cookies(strUniqueID & "User")("Name") & "'; "

					my_Conn.Execute(strSql)			
					call NTauthenticate()
					call NTUser()	
				end if
			end if
		end if
		if (Request.Form("strAuthType") <> strAuthType) and strAuthType = "nt" then 
			mLev = cint(ChkUser2(Request.Cookies(strUniqueID & "User")("Name"), Request.Cookies(strUniqueID & "User")("Pword")))
			if not(mLev = 4 and getMemberNumber(Request.Cookies(strUniqueID & "User")("Name")) = 1) then
				Err_Msg = Err_Msg & "<li>ֻ�й���Ա�����޸���̳����Ȩģʽ</li>"
			else
				Session(strCookieURL & "Approval") = "" 
			end if	
		end if
		if Err_Msg = "" then

			

			'## Forum_SQL
			strSql = "UPDATE " & strTablePrefix & "CONFIG "
			strSql = strSql & " SET C_STRFORUMTITLE = '" & ChkString(Request.Form("strForumTitle"),"title") & "'"
			strSql = strSql & ",    C_STRCOPYRIGHT = '" & ChkString(Request.Form("strCopyright"),"title") & "'"
			strSql = strSql & ",    C_STRTITLEIMAGE = '" & ChkString(Request.Form("strTitleImage"),"url") & "'"
			strSql = strSql & ",    C_STRHOMEURL = '" & ChkString(Request.Form("strHomeURL"),"url") & "'"
			strSql = strSql & ",    C_STRFORUMURL = '" & ChkString(Request.Form("strForumURL"),"url") & "'"
			strSql = strSql & ",    C_STRAUTHTYPE = '" & Request.Form("strAuthType") & "'"
			strSql = strSql & ",    C_STRSETCOOKIETOFORUM = " & Request.Form("strSetCookieToForum")
			strSql = strSql & ",    C_STRGFXBUTTONS = " & Request.Form("strGfxButtons")
			strSql = strSql & ",    C_STRSHOWIMAGEPOWEREDBY = " & Request.Form("strShowImagePoweredBy")
			strSql = strSql & " WHERE CONFIG_ID = " & 1
			my_Conn.Execute (strSql)
			Application(strCookieURL & "ConfigLoaded") = ""

%>
<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">��̳�����������</font></p>
<meta http-equiv="Refresh" content="0; URL=admin_home.asp">

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">�޸���ɣ�</font></p>

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="admin_home.asp">�ص���������</font></a></p>
<%		else %>
	<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">������������������û����ȫ����</font></p>

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
<form action="admin_config_system.asp" method="post" id="Form1" name="Form1">
<input type="hidden" name="Method_Type" value="Write_Configuration">
<table border="0" cellspacing="1" cellpadding="4" align=center bgcolor="<% =strPopUpBorderColor %>">
  <tr valign="top">
    <td bgcolor="<% =strHeadCellColor %>" colspan="2"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>" color="<% =strHeadFontColor %>">��Ҫ�趨</font></td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">��̳���ƣ�</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    <input type="text" name="strForumTitle" size="25" value="<% if strForumTitle <> "" then Response.Write(strForumTitle) else '## do nothing %>">
    <a href="JavaScript:openWindow3('pop_config_help.asp#forumtitle')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">��̳��Ȩ˵����</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    <input type="text" name="strCopyright" size="25" value="<% if strCopyright <> "" then Response.Write(strCopyright) else Response.Write("&copy; 2000 Snitz Communications") %>">
    <a href="JavaScript:openWindow3('pop_config_help.asp#copyright')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">��̳logoͼƬ��ַ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    <input type="text" name="strTitleImage" size="25" value="<% if strTitleImage <> "" then Response.Write(strTitleImage) else Response.Write("logo_snitz_forums_2000.gif") %>">
    <a href="JavaScript:openWindow3('pop_config_help.asp#titleimage')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">��վ��ҳ��</font></td></td>
    <td bgColor="<% =strPopUpTableColor %>">
    <input type="text" name="strHomeURL" size="25" value="<% if strHomeURL <> "" then Response.Write(strHomeURL) else '## Do Nothing %>">
    <a href="JavaScript:openWindow3('pop_config_help.asp#homeurl')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">��̳��ַ��</font></td></td>
    <td bgColor="<% =strPopUpTableColor %>">
    <input type="text" name="strForumURL" size="25" value="<% if strForumURL <> "" then Response.Write(strForumURL) else '## do nothing %>">
    <a href="JavaScript:openWindow3('pop_config_help.asp#forumurl')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�汾ѶϢ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% if strVersion <> "" then Response.Write("[<b>"& strVersion & "</b>]") else Response.Write("<b>[Couldn't read version info..]</b>") %></font>
    <!--<a href="JavaScript:openWindow3('pop_config_help.asp#forumtitle')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>-->
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">��Ȩģʽ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    DB: <input type="radio" class="radio" name="strAuthType" value="db" <% if strAuthType = "db" then Response.Write("checked") %>> 
    NT: <input type="radio" class="radio" name="strAuthType" value="nt" <% if strAuthType = "nt" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#AuthType')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </font></td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�趨 Cookie ����</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    ��������<input type="radio" class="radio" name="strSetCookieToForum" value="1" <% if strSetCookieToForum = "1" then Response.Write("checked") %>> 
    ��վ��<input type="radio" class="radio" name="strSetCookieToForum" value="0" <% if strSetCookieToForum = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#SetCookieToForum')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </font></td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">ʹ��ͼ�ΰ�ť��</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    ������<input type="radio" class="radio" name="strGfxButtons" value="1" <% if strGfxButtons = "1" then Response.Write("checked") %>> 
    �رգ�<input type="radio" class="radio" name="strGfxButtons" value="0" <% if strGfxButtons = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#GfxButtons')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </font></td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">��Ȩ��ʾ��ʽ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    ͼ�Σ�<input type="radio" class="radio" name="strShowImagePoweredBy" value="1" <% if strShowImagePoweredBy = "1" then Response.Write("checked") %>> 
    ���֣�<input type="radio" class="radio" name="strShowImagePoweredBy" value="0" <% if strShowImagePoweredBy = "0" then Response.Write("checked") %>>
<!--    <a href="JavaScript:openWindow3('pop_config_help.asp#GfxButtons')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>-->
    </font></td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" colspan="2" align="center"><input type="submit" value="�ύ�趨" id="submit1" name="submit1"> <input type="reset" value="��������" id="reset1" name="reset1"></td>
  </tr>
</table>
</form>
</font>
<%	end if %>
<!--#INCLUDE file="inc_footer.asp" -->
<% Else %>
<% Response.Redirect "admin_login.asp" %>
<% End IF %>
