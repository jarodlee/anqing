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
    <img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="default.asp">������̳��ҳ</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="admin_home.asp">��̳��������</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" height=15 width=15 border="0">&nbsp;��Աע��ѡ������<br>
    </font></td>
  </tr>
</table>
<%
	if Request.Form("Method_Type") = "Write_Configuration" then 
		Err_Msg = ""
	
		if Err_Msg = "" then

			'## Forum_SQL
			strSql = "UPDATE " & strTablePrefix & "CONFIG "
			strSql = strSql & " SET C_STRFULLNAME   = " & Request.Form("strFullName") & " "
			strSql = strSql & ",    C_STRPICTURE    = " & Request.Form("strPicture") & " "
			strSql = strSql & ",    C_STRSEX        = " & Request.Form("strSex") & " "
			strSql = strSql & ",    C_STRCITY       = " & Request.Form("strCity") & " "
			strSql = strSql & ",    C_STRSTATE      = " & Request.Form("strState") & " "
			strSql = strSql & ",    C_STRAGE        = " & Request.Form("strAge") & " "
			strSql = strSql & ",    C_STRCOUNTRY    = " & Request.Form("strCountry") & " "
			strSql = strSql & ",    C_STROCCUPATION = " & Request.Form("strOccupation") & " "
			strSql = strSql & ",    C_STRHOMEPAGE   = " & Request.Form("strHomepage") & " "
			strSql = strSql & ",    C_STRFAVLINKS   = " & Request.Form("strFavLinks") & " "
			strSql = strSql & ",    C_STRICQ        = " & Request.Form("strICQ") & " "
			strSql = strSql & ",    C_STRYAHOO      = " & Request.Form("strYAHOO") & " "
			strSql = strSql & ",    C_STRAIM        = " & Request.Form("strAIM") & " "
			strSql = strSql & ",    C_STRBIO        = " & Request.Form("strBio") & " "
			strSql = strSql & ",    C_STRHOBBIES	= " & Request.Form("strHobbies") & " "
			strSql = strSql & ",    C_STRLNEWS      = " & Request.Form("strLNews") & " "
			strSql = strSql & ",    C_STRQUOTE		= " & Request.Form("strQuote") & " "
			strSql = strSql & ",    C_STRMARSTATUS  = " & Request.Form("strMarStatus") & " "
			strSql = strSql & ",    C_STRRECENTTOPICS = " & Request.Form("strRecentTopics") & " "
			strSql = strSql & " WHERE CONFIG_ID = " & 1


			my_Conn.Execute (strSql)
			Application(strCookieURL & "ConfigLoaded") = ""

%>
<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">��Աע��ѡ�������Ѿ��������</font></p>
<meta http-equiv="Refresh" content="0; URL=admin_home.asp">

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">�޸���ɣ�</font></p>

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="admin_home.asp">������̳��������</font></a></p>
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
<form action="admin_config_members.asp" method="post" id="Form1" name="Form1">
<input type="hidden" name="Method_Type" value="Write_Configuration">
<table border="0" cellspacing="1" cellpadding="4" align=center bgcolor="<% =strPopUpBorderColor %>">
  <tr valign="top">
    <td bgcolor="<% =strHeadCellColor %>" colspan="2"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">��Աע��ѡ������</font></td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">��ʵ������</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    ������<input type="radio" class="radio" name="strFullName" value="1"<% if strFullName <> "0" then Response.Write(" checked") %>> 
    �رգ�<input type="radio" class="radio" name="strFullName" value="0"<% if strFullName = "0" then Response.Write(" checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#FullName')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">����ͷ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    ������<input type="radio" class="radio" name="strPicture" value="1"<% if strPicture <> "0" then Response.Write(" checked") %>> 
    �رգ�<input type="radio" class="radio" name="strPicture" value="0"<% if strPicture = "0" then Response.Write(" checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#Picture')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">��󷢱�</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    ������<input type="radio" class="radio" name="strRecentTopics" value="1" <% if strRecentTopics <> "0" then Response.Write("checked") %>> 
    �رգ�<input type="radio" class="radio" name="strRecentTopics" value="0" <% if strRecentTopics = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#RecentTopics')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">��&nbsp;&nbsp;&nbsp;&nbsp;��</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    ������<input type="radio" class="radio" name="strSex" value="1"<% if strSex <> "0" then Response.Write(" checked") %>> 
    �رգ�<input type="radio" class="radio" name="strSex" value="0"<% if strSex = "0" then Response.Write(" checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#Sex')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
    <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">��&nbsp;&nbsp;&nbsp;&nbsp;�䣺</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    ������<input type="radio" class="radio" name="strAge" value="1"<% if strAge <> "0" then Response.Write(" checked") %>> 
    �رգ�<input type="radio" class="radio" name="strAge" value="0"<% if strAge = "0" then Response.Write(" checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#Sex')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">��ס���У�</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    ������<input type="radio" class="radio" name="strCity" value="1"<% if strCity <> "0" then Response.Write(" checked") %>> 
    �رգ�<input type="radio" class="radio" name="strCity" value="0"<% if strCity = "0" then Response.Write(" checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#City')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">����ʡ�֣�</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    ������<input type="radio" class="radio" name="strState" value="1"<% if strState <> "0" then Response.Write(" checked") %>> 
    �رգ�<input type="radio" class="radio" name="strState" value="0"<% if strState = "0" then Response.Write(" checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#State')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">��&nbsp;&nbsp;&nbsp;&nbsp;�ң�</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    ������<input type="radio" class="radio" name="strCountry" value="1" <% if (lcase(strCountry) <> "0") then Response.Write("checked")%>> 
    �رգ�<input type="radio" class="radio" name="strCountry" value="0" <% if (lcase(strCountry) = "0") then Response.Write("checked")%>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#Country')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
   </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">ICQ ���룺</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    ������<input type="radio" class="radio" name="strICQ" value="1" <% if strICQ <> "0" then Response.Write("checked") %>> 
    �رգ�<input type="radio" class="radio" name="strICQ" value="0" <% if strICQ = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#icq')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">OICQ���룺</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    ������<input type="radio" class="radio" name="strYAHOO" value="1" <% if strYAHOO <> "0" then Response.Write("checked") %>> 
    �رգ�<input type="radio" class="radio" name="strYAHOO" value="0" <% if strYAHOO = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#yahoo')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">AIM���룺</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    ������<input type="radio" class="radio" name="strAIM" value="1" <% if strAIM <> "0" then Response.Write("checked") %>> 
    �رգ�<input type="radio" class="radio" name="strAIM" value="0" <% if strAIM = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#aim')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">����ְҵ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    ������<input type="radio" class="radio" name="strOccupation" value="1"<% if strOccupation <> "0" then Response.Write(" checked") %>> 
    �رգ�<input type="radio" class="radio" name="strOccupation" value="0"<% if strOccupation = "0" then Response.Write(" checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#Occupation')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">������ҳ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    ������<input type="radio" class="radio" name="strHomepage" value="1" <% if strHomepage <> "0" then Response.Write("checked") %>> 
    �رգ�<input type="radio" class="radio" name="strHomepage" value="0" <% if strHomepage = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#homepages')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�Ƽ���ҳ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    ������<input type="radio" class="radio" name="strFavLinks" value="1" <% if strFavLinks <> "0" then Response.Write("checked") %>> 
    �رգ�<input type="radio" class="radio" name="strFavLinks" value="0" <% if strFavLinks = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#FavLinks')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">����״����</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    ������<input type="radio" class="radio" name="strMarStatus" value="1" <% if strMarStatus <> "0" then Response.Write("checked") %>> 
    �رգ�<input type="radio" class="radio" name="strMarStatus" value="0" <% if strMarStatus = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#MStatus')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">���˼�飺</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    ������<input type="radio" class="radio" name="strBio" value="1" <% if strBio <> "0" then Response.Write("checked") %>> 
    �رգ�<input type="radio" class="radio" name="strBio" value="0" <% if strBio = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#Bio')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
   <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">����ϲ�ã�</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    ������<input type="radio" class="radio" name="strHobbies" value="1" <% if strHobbies <> "0" then Response.Write("checked") %>> 
    �رգ�<input type="radio" class="radio" name="strHobbies" value="0" <% if strHobbies = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#hobbies')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">���״����</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    ������<input type="radio" class="radio" name="strLNews" value="1" <% if strLNews <> "0" then Response.Write("checked") %>> 
    �رգ�<input type="radio" class="radio" name="strLNews" value="0" <% if strLNews = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#LNews')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�� �� ����</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    ������<input type="radio" class="radio" name="strQuote" value="1" <% if strQuote <> "0" then Response.Write("checked") %>> 
    �رգ�<input type="radio" class="radio" name="strQuote" value="0" <% if strQuote = "0" then Response.Write("checked") %>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#Quote')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
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
