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
    <img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="default.asp">������̳��ҳ</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="admin_home.asp">��̳��������</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" height=15 width=15 border="0">&nbsp;��̳����<br>
    </font></td>
  </tr>
</table>
<%
	if Request.Form("Method_Type") = "Write_Configuration" then 
		Err_Msg = ""
		if Request.Form("strIMGInPosts") = "1" and Request.Form("strAllowForumCode") = "0" then 
			Err_Msg = Err_Msg & "<li>���������ͼ�ͱ��뿪������ר�ô���</li>"
		end if
		if (Request.Form("strHotTopic") = "1" and strHotTopic = "1") or (Request.Form("strHotTopic") = "1" and strHotTopic = "0") then
			if Request.Form("intHotTopicNum") = "" then 
				Err_Msg = Err_Msg & "<li>������趨���������������</li>"
			end if
			if left(Request.Form("intHotTopicNum"), 1) = "-" then 
				Err_Msg = Err_Msg & "<li>������趨ȷ�е���������������</li>"
			end if
			if left(Request.Form("intHotTopicNum"), 1) = "+" then 
				Err_Msg = Err_Msg & "<li>ֻ����������ֲ�Ҫ���� <b>+</b></li>"
			end if
		end if

		if Err_Msg = "" then

			'## Forum_SQL
			strSql = "UPDATE " & strTablePrefix & "CONFIG "
			strSql = strSql & " SET C_STRMOVETOPICMODE         = " & Request.Form("strMoveTopicMode") & ""
			strSql = strSql & ",    C_STRIPLOGGING             = " & Request.Form("strIPLogging") & ""
			strSql = strSql & ",    C_STRPRIVATEFORUMS         = " & Request.Form("strPrivateForums") & ""
			strSql = strSql & ",    C_STRSHOWMODERATORS        = " & Request.Form("strShowModerators") & ""
			strSql = strSql & ",    C_STRALLOWFORUMCODE        = " & Request.Form("strAllowForumCode") & ""
			strSql = strSql & ",    C_STRIMGINPOSTS            = " & Request.Form("strIMGInPosts") & ""
			strSql = strSql & ",    C_STRALLOWHTML             = " & Request.Form("strAllowHTML") & ""
			strSql = strSql & ",    C_STRHOTTOPIC              = " & Request.Form("strHotTopic") & ""
			if (Request.Form("strHotTopic") = "1" and strHotTopic = "1") or (Request.Form("strHotTopic") = "1" and strHotTopic = "0") then
				strSql = strSql & ",    C_INTHOTTOPICNUM           = " & Request.Form("intHotTopicNum") & ""
			end if
			strSql = strSql & ",    C_STRiconS                 = " & Request.Form("stricons") & ""
			strSql = strSql & ",    C_STRSECUREADMIN           = " & Request.Form("strSecureAdmin") & ""
			strSql = strSql & ",    C_STRNOCOOKIES             = " & Request.Form("strNoCookies") & ""
			strSql = strSql & ",    C_STREDITEDBYDATE          = " & Request.Form("strEditedByDate") & ""
			strSql = strSql & ",    C_STRSHOWSTATISTICS        = " & Request.Form("strShowStatistics") & ""
			strSql = strSql & ",    C_STRSHOWPAGING            = " & Request.Form("strShowPaging") & ""
			strSql = strSql & ",    C_STRSHOWTOPICNAV          = " & Request.Form("strShowTopicNav") & ""
			strSql = strSql & ",    C_STRPAGESIZE			   = " & Request.Form("strPageSize") & ""
			strSql = strSql & ",    C_STRPAGENUMBERSIZE        = " & Request.Form("strPageNumberSize") & ""
			strSql = strSql & " WHERE CONFIG_ID = " & 1

			my_Conn.Execute (strSql)
			Application(strCookieURL & "ConfigLoaded") = ""

%>
<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">��̳�������</font></p>
<meta http-equiv="Refresh" content="0; URL=admin_home.asp">

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
<form action="admin_config_features.asp" method="post" id="Form1" name="Form1">
<input type="hidden" name="Method_Type" value="Write_Configuration">
<table border="0" cellspacing="1" cellpadding="3" align=center bgcolor="<% =strPopUpBorderColor %>">
  <tr valign="top">
    <td bgcolor="<% =strHeadCellColor %>" colspan="2"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">��̳����</font></td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">���ư����ƶ��ð����⣺</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    ������<input type="radio" class="radio" name="strMoveTopicMode" value="1"<% if strMoveTopicMode <> "0" then Response.Write(" checked") %>> 
    �رգ�<input type="radio" class="radio" name="strMoveTopicMode" value="0"<% if strMoveTopicMode = "0" then Response.Write(" checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#MoveTopicMode')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">��¼IP��</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    ������<input type="radio" class="radio" name="strIPLogging" value="1"<% if strIPLogging <> "0" then Response.Write(" checked") %>> 
    �رգ�<input type="radio" class="radio" name="strIPLogging" value="0"<% if strIPLogging = "0" then Response.Write(" checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#IPLogging')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">������̳��</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    ������<input type="radio" class="radio" name="strPrivateForums" value="1"<% if strPrivateForums <> "0" then Response.Write(" checked") %>> 
    �رգ�<input type="radio" class="radio" name="strPrivateForums" value="0"<% if strPrivateForums = "0" then Response.Write(" checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#privateforums')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">��ʾ������</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    ������<input type="radio" class="radio" name="strShowModerators" value="1"<% if strShowModerators <> "0" then Response.Write(" checked") %>> 
    �رգ�<input type="radio" class="radio" name="strShowModerators" value="0"<% if strShowModerators = "0" then Response.Write(" checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#ShowModerator')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">����ר�ô��룺&nbsp;</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    ������<input type="radio" class="radio" name="strAllowForumCode" value="1"<% if strAllowForumCode <> "0" then Response.Write(" checked") %>> 
    �رգ�<input type="radio" class="radio" name="strAllowForumCode" value="0"<% if strAllowForumCode = "0" then Response.Write(" checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#AllowForumCode')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">��ͼ���ܣ�</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    ������<input type="radio" class="radio" name="strIMGInPosts" value="1" <% if (lcase(strIMGInPosts) <> "0") then Response.Write("checked")%>> 
    �رգ�<input type="radio" class="radio" name="strIMGInPosts" value="0" <% if (lcase(strIMGInPosts) = "0") then Response.Write("checked")%>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#imginposts')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
   </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">����HTML���룺</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    ������<input type="radio" class="radio" name="strAllowHTML" value="1"<% if strAllowHTML <> "0" then Response.Write(" checked") %>> 
    �رգ�<input type="radio" class="radio" name="strAllowHTML" value="0"<% if strAllowHTML = "0" then Response.Write(" checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#AllowHTML')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">��ȫ����ģʽ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    ������<input type="radio" class="radio" name="strSecureAdmin" value="1" <% if lcase(strSecureAdmin) <> "0" then Response.Write("checked") %>> 
    �رգ�<input type="radio" class="radio" name="strSecureAdmin" value="0" <% if lcase(strSecureAdmin) = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#secureadminmode')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�� Cookie ģʽ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    ������<input type="radio" class="radio" name="strNoCookies" value="1" <% if lcase(strNoCookies) <> "0" then Response.Write("checked") %>> 
    �رգ�<input type="radio" class="radio" name="strNoCookies" value="0" <% if lcase(strNoCookies) = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#allownoncookies')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">��ʾ�ر����ڣ�</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    ������<input type="radio" class="radio" name="strEditedByDate" value="1" <% if lcase(strEditedByDate) <> "0" then Response.Write("checked") %>> 
    �رգ�<input type="radio" class="radio" name="strEditedByDate" value="0" <% if lcase(strEditedByDate) = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#editedbydate')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�������⣺</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    ������<input type="radio" class="radio" name="strHotTopic" value="1" <% if (strHotTopic <> "0" or lcase(HotTopic) <> "0") then Response.Write("checked") %>> 
    �رգ�<input type="radio" class="radio" name="strHotTopic" value="0" <% if (strHotTopic = "0" or lcase(HotTopic) = "0") then Response.Write("checked") %>>
    <input type="text" name="intHotTopicNum" size="5" value="<% if intHotTopicNum <> "" then Response.Write(intHotTopicNum) else Response.Write("20") %>">
    <a href="JavaScript:openWindow3('pop_config_help.asp#hottopics')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">����ת��ͼʾ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    ������<input type="radio" class="radio" name="stricons" value="1" <% if (lcase(stricons) <> "0" or lcase(Smiles) <> "0") then Response.Write("checked") %>> 
    �رգ�<input type="radio" class="radio" name="stricons" value="0" <% if (lcase(stricons) = "0" or lcase(Smiles) = "0") then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#<%=strImageURL %>icons')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
   <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">��ϸ״̬�У�</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    ������<input type="radio" class="radio" name="strShowStatistics" value="1" <% if strShowStatistics <> "0" then Response.Write("checked") %>> 
    �رգ�<input type="radio" class="radio" name="strShowStatistics" value="0" <% if strShowStatistics = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#stats')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">ҳ��������ӣ�</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    ������<input type="radio" class="radio" name="strShowPaging" value="1" <% if strShowPaging <> "0" then Response.Write("checked") %>> 
    �رգ�<input type="radio" class="radio" name="strShowPaging" value="0" <% if strShowPaging = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#ShowPaging')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">��ʾ ǰ/�� ���⣺</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    ������<input type="radio" class="radio" name="strShowTopicNav" value="1" <% if strShowTopicNav <> "0" then Response.Write("checked") %>> 
    �رգ�<input type="radio" class="radio" name="strShowTopicNav" value="0" <% if strShowTopicNav = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#ShowTopicNav')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">ÿҳ��ʾ��������</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    <input type="text" name="strPageSize" size="5" value="<% if strPageSize <> "" then Response.Write(strPageSize) else Response.Write("15") %>">
<!--    <a href="JavaScript:openWindow3('pop_config_help.asp#hottopics')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>-->
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">ÿ����ʾ��������</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    <input type="text" name="strPageNumberSize" size="5" value="<% if strPageNumberSize <> "" then Response.Write(strPageNumberSize) else Response.Write("10") %>">
<!--    <a href="JavaScript:openWindow3('pop_config_help.asp#hottopics')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>-->
    </td>
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
