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
    <img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="default.asp">������̳��ҳ</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="admin_home.asp">��̳��������</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" height=15 width=15 border="0">&nbsp;�������<br>
    </font></td>
  </tr>
</table>
<%
	if Request.Form("Method_Type") = "Write_Configuration" then 

		'## Forum_SQL
		strSql = "UPDATE " & strTablePrefix & "CONFIG "
		strSql = strSql & " SET C_STRDEFAULTFONTFACE       = '" & Request("strDefaultFontFace") & "', "
		strSql = strSql & "     C_STRDEFAULTFONTSIZE       = '" & Request.Form("strDefaultFontSize") & "', "
		strSql = strSql & "     C_STRHEADERFONTSIZE        = '" & Request.Form("strHeaderFontSize") & "', "
		strSql = strSql & "     C_STRFOOTERFONTSIZE        = '" & Request.Form("strFooterFontSize") & "', "
		strSql = strSql & "     C_STRPAGEBGIMAGE           = '" & Request.Form("strPageBGImage") & "', "
		strSql = strSql & "     C_STRPAGEBGCOLOR           = '" & Request.Form("strPageBGColor") & "', "
		strSql = strSql & "     C_STRDEFAULTFONTCOLOR      = '" & Request.Form("strDefaultFontColor") & "', "
		strSql = strSql & "     C_STRLINKCOLOR             = '" & Request.Form("strLinkColor") & "', "
		strSql = strSql & "     C_STRLINKTEXTDECORATION    = '" & Request.Form("strLinkTextDecoration") & "', "
		strSql = strSql & "     C_STRVISITEDLINKCOLOR      = '" & Request.Form("strVisitedLinkColor") & "', "
		strSql = strSql & "     C_STRVISITEDTEXTDECORATION = '" & Request.Form("strVisitedTextDecoration") & "', "
		strSql = strSql & "     C_STRACTIVELINKCOLOR       = '" & Request.Form("strActiveLinkColor") & "', "
		strSql = strSql & "     C_STRHOVERFONTCOLOR        = '" & Request.Form("strHoverFontColor") & "', "
		strSql = strSql & "     C_STRHOVERTEXTDECORATION   = '" & Request.Form("strHoverTextDecoration") & "', "
		strSql = strSql & "     C_STRHEADCELLCOLOR         = '" & Request.Form("strHeadCellColor") & "', "
		strSql = strSql & "     C_STRHEADFONTCOLOR         = '" & Request.Form("strHeadFontColor") & "', "
		strSql = strSql & "     C_STRCATEGORYCELLCOLOR     = '" & Request.Form("strCategoryCellColor") & "', "
		strSql = strSql & "     C_STRCATEGORYFONTCOLOR     = '" & Request.Form("strCategoryFontColor") & "', "
		strSql = strSql & "     C_STRFORUMFIRSTCELLCOLOR   = '" & Request.Form("strForumFirstCellColor") & "', "
		strSql = strSql & "     C_STRFORUMCELLCOLOR        = '" & Request.Form("strForumCellColor") & "', "
		strSql = strSql & "     C_STRALTFORUMCELLCOLOR     = '" & Request.Form("strAltForumCellColor") & "', "
		strSql = strSql & "     C_STRFORUMFONTCOLOR        = '" & Request.Form("strForumFontColor") & "', "
		strSql = strSql & "     C_STRFORUMLINKCOLOR        = '" & Request.Form("strForumLinkColor") & "', "
		strSql = strSql & "     C_STRTABLEBORDERCOLOR      = '" & Request.Form("strTableBorderColor") & "', "
		strSql = strSql & "     C_STRPOPUPTABLECOLOR       = '" & Request.Form("strPopUpTableColor") & "', "
		strSql = strSql & "     C_STRPOPUPBORDERCOLOR      = '" & Request.Form("strPopUpBorderColor") & "', "
		strSql = strSql & "     C_STRNEWFONTCOLOR          = '" & Request.Form("strNewFontColor") & "', "
		strSql = strSql & "     C_STRTOPICWIDTHLEFT        = '" & Request.Form("strTopicWidthLeft") & "', "
		strSql = strSql & "     C_STRTOPICNOWRAPLEFT       = " & Request.Form("strTopicNoWrapLeft") & ", "
		strSql = strSql & "     C_STRTOPICWIDTHRIGHT       = '" & Request.Form("strTopicWidthRight") & "', "
		strSql = strSql & "     C_STRTOPICNOWRAPRIGHT      = " & Request.Form("strTopicNoWrapRight") & " "
		strSql = strSql & " WHERE CONFIG_ID = 1"
		
'		Response.Write(strSql)
'		Response.End

		my_Conn.Execute (strSql)
		Application(strCookieURL & "ConfigLoaded") = ""

%>
<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">��������Ѿ�������ϣ�</font></p>
<meta http-equiv="Refresh" content="0; URL=admin_home.asp">

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">�޸���ɣ�</font></p>

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="admin_home.asp">���ع�������</font></a></p>
<%	else %>
<form method="post" action="colortemplates.asp">
<table border="0" cellspacing="1" cellpadding="4" align=center bgcolor="<% =strPopUpBorderColor %>">
<% '########################################## %>
  <tr valign="top"> 
  <td bgcolor="<% =strHeadCellColor %>" colspan="2"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">��̳��ɫ����</font></td>
  </tr>
  <tr align="center" valign="top"> 
  <td bgcolor="<% =strPopUpTableColor %>" align="right" width="70%"> 
  <div align="center"> 
    <div align="Right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�����������ѡ��һ����̳�ṩ����ɫ����</font></div>
    <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"> </font></div>
  </td>
  <td bgcolor="<% =strPopUpTableColor %>" width="30%"> 
  <div align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"> 
    <select name="select">
      <option value="default">Ĭ��ֵ</option>
      <option value="beige">��ɫ</option>
      <option value="BPISG">����</option>
      <option value="darkstation">Dark Station</option>
      <option value="green">Green</option>
      <option value="grey">Grey</option>
      <option value="redrose">Red Rose</option>
      <option value="slide">Slide</option>
      <option value="fantasy">Fantasy</option>
      <option value="px">Px</option>
      <option value="madonna">Madonna</option>
      <option value="blacksun">Black Sun</option>
    </select>
    </font></div>
  </td>
  </tr>
  <tr align="center" valign="top"> 
  <td bgcolor="<% =strPopUpTableColor %>" colspan="2"><font face="courier" size="<% =strDefaultFontSize %>"> 
  <input type="hidden" name="template" value="Write">
  <input type="submit" name="Submit" value="ʹ�ô���ɫ����">
  </font></td>
  </tr>
  </table>
</form>
<form action="admin_config_colors.asp" method="post" id="Form1" name="Form1">
<input type="hidden" name="Method_Type" value="Write_Configuration">
<table border="0" cellspacing="1" cellpadding="4" align=center bgcolor="<% =strPopUpBorderColor %>">
<% '################################## %>
  <tr valign="top">
    <td bgcolor="<% =strHeadCellColor %>" colspan="2"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">������ʽ/����С/�����ɫ�趨</font></td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">���壺</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><input type="text" name="strDefaultFontFace" size="25" value="<% if strDefaultFontFace <> "" then Response.Write(strDefaultFontFace) else '## do nothing %>">
    <a href="JavaScript:openWindow3('pop_config_help.asp#fontfacetype')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">Ĭ�������С��</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    <select name="strDefaultFontSize">
      <option<% if (lcase(strDefaultFontSize)="1") then Response.Write(" selected") %> value="1">1 (8 pt)</option>
      <option<% if (lcase(strDefaultFontSize)="2") then Response.Write(" selected") %> value="2">2 (10 pt)</option>
      <option<% if (lcase(strDefaultFontSize)="3") then Response.Write(" selected") %> value="3">3 (12 pt)</option>
      <option<% if (lcase(strDefaultFontSize)="4") then Response.Write(" selected") %> value="4">4 (14 pt)</option>
      <option<% if (lcase(strDefaultFontSize)="5") then Response.Write(" selected") %> value="5">5 (18 pt)</option>
      <option<% if (lcase(strDefaultFontSize)="6") then Response.Write(" selected") %> value="6">6 (24 pt)</option>
      <option<% if (lcase(strDefaultFontSize)="6") then Response.Write(" selected") %> value="7">7 (36 pt)</option>
    </select>
    <a href="JavaScript:openWindow3('pop_config_help.asp#fontsize')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">���������С��</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    <select name="strHeaderFontSize">
      <option<% if (lcase(strHeaderFontSize)="1") then Response.Write(" selected") %> value="1">1 (8 pt)</option>
      <option<% if (lcase(strHeaderFontSize)="2") then Response.Write(" selected") %> value="2">2 (10 pt)</option>
      <option<% if (lcase(strHeaderFontSize)="3") then Response.Write(" selected") %> value="3">3 (12 pt)</option>
      <option<% if (lcase(strHeaderFontSize)="4") then Response.Write(" selected") %> value="4">4 (14 pt)</option>
      <option<% if (lcase(strHeaderFontSize)="5") then Response.Write(" selected") %> value="5">5 (18 pt)</option>
      <option<% if (lcase(strHeaderFontSize)="6") then Response.Write(" selected") %> value="6">6 (24 pt)</option>
      <option<% if (lcase(strHeaderFontSize)="6") then Response.Write(" selected") %> value="7">7 (36 pt)</option>
    </select>
    <a href="JavaScript:openWindow3('pop_config_help.asp#fontsize')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">ҳü�����С��</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    <select name="strFooterFontSize">
      <option<% if (lcase(strFooterFontSize)="1") then Response.Write(" selected") %> value="1">1 (8 pt)</option>
      <option<% if (lcase(strFooterFontSize)="2") then Response.Write(" selected") %> value="2">2 (10 pt)</option>
      <option<% if (lcase(strFooterFontSize)="3") then Response.Write(" selected") %> value="3">3 (12 pt)</option>
      <option<% if (lcase(strFooterFontSize)="4") then Response.Write(" selected") %> value="4">4 (14 pt)</option>
      <option<% if (lcase(strFooterFontSize)="5") then Response.Write(" selected") %> value="5">5 (18 pt)</option>
      <option<% if (lcase(strFooterFontSize)="6") then Response.Write(" selected") %> value="6">6 (24 pt)</option>
      <option<% if (lcase(strFooterFontSize)="6") then Response.Write(" selected") %> value="7">7 (36 pt)</option>
    </select>
    <a href="JavaScript:openWindow3('pop_config_help.asp#fontsize')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">ҳ�汳��ͼ����</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><input type="text" name="strPageBGImage" size="20" value="<% if strPageBGImage <> "" then Response.Write(strPageBGImage) else '## do nothing%>">
    </td>
  </tr>
  
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">������ɫ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><input type="text" name="strPageBGColor" size="10" value="<% if strPageBGColor <> "" then Response.Write(strPageBGColor) else '## do nothing%>">
    <a href="JavaScript:openWindow3('pop_config_help.asp#colors')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">������ɫ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><input type="text" name="strDefaultFontColor" size="10" value="<% if strDefaultFontColor <> "" then Response.Write(strDefaultFontColor) else '## do nothing %>">
    <a href="JavaScript:openWindow3('pop_config_help.asp#colors')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">������ɫ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><input type="text" name="strLinkColor" size="10" value="<% if strLinkColor <> "" then Response.Write(strLinkColor) else '## do nothing %>">
    <a href="JavaScript:openWindow3('pop_config_help.asp#colors')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">������ʽ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    <select name="strLinkTextDecoration">
      <option<% if (lcase(strLinkTextDecoration)="none" or lcase(LinkTextDecoration)="none") then Response.Write(" selected") %>>��</option>
      <option<% if (lcase(strLinkTextDecoration)="blink" or lcase(LinkTextDecoration)="blink") then Response.Write(" selected") %>>��˸</option>
      <option<% if (lcase(strLinkTextDecoration)="line-through" or lcase(LinkTextDecoration)="line-through") then Response.Write(" selected") %>>ɾ����</option>
      <option<% if (lcase(strLinkTextDecoration)="overline" or lcase(LinkTextDecoration)="overline") then Response.Write(" selected") %>>����</option>
      <option<% if (lcase(strLinkTextDecoration)="underline" or lcase(LinkTextDecoration)="underline") then Response.Write(" selected") %>>����</option>
    </select>
    <a href="JavaScript:openWindow3('pop_config_help.asp#fontdecorations')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </font></td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�ѷ���������ɫ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><input type="text" name="strVisitedLinkColor" size="10" value="<% if strVisitedLinkColor <> "" then Response.Write(strVisitedLinkColor) else '## do nothing %>">
    <a href="JavaScript:openWindow3('pop_config_help.asp#colors')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�ѷ���������ʽ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    <select name="strVisitedTextDecoration">
      <option<% if (lcase(strVisitedTextDecoration)="none" or lcase(VisitedTextDecoration)="none") then Response.Write(" selected") %>>��</option>
      <option<% if (lcase(strVisitedTextDecoration)="blink" or lcase(VisitedTextDecoration)="blink") then Response.Write(" selected") %>>��˸</option>
      <option<% if (lcase(strVisitedTextDecoration)="line-through" or lcase(VisitedTextDecoration)="line-through") then Response.Write(" selected") %>>ɾ����</option>
      <option<% if (lcase(strVisitedTextDecoration)="overline" or lcase(VisitedTextDecoration)="overline") then Response.Write(" selected") %>>����</option>
      <option<% if (lcase(strVisitedTextDecoration)="underline" or lcase(VisitedTextDecoration)="underline") then Response.Write(" selected") %>>����</option>
    </select>
    <a href="JavaScript:openWindow3('pop_config_help.asp#fontdecorations')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">���ͣ�������ϵ���ɫ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><input type="text" name="strActiveLinkColor" size="10" value="<% if strActiveLinkColor <> "" then Response.Write(strActiveLinkColor) else '## do nothing %>">
    <a href="JavaScript:openWindow3('pop_config_help.asp#colors')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">���ͣ������ʱ������ɫ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><input type="text" name="strHoverFontColor" size="10" value="<% if strHoverFontColor <> "" then Response.Write(strHoverFontColor) else '## do nothing %>">
    <a href="JavaScript:openWindow3('pop_config_help.asp#colors')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">���ͣ������ʱ������ʽ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    <select name="strHoverTextDecoration">
      <option<% if (lcase(strHoverTextDecoration)="none" or lcase(HoverTextDecoration)="none") then Response.Write(" selected") %>>��</option>
      <option<% if (lcase(strHoverTextDecoration)="blink" or lcase(HoverTextDecoration)="blink") then Response.Write(" selected") %>>��˸</option>
      <option<% if (lcase(strHoverTextDecoration)="line-through" or lcase(HoverTextDecoration)="line-through") then Response.Write(" selected") %>>ɾ����</option>
      <option<% if (lcase(strHoverTextDecoration)="overline" or lcase(HoverTextDecoration)="overline") then Response.Write(" selected") %>>����</option>
      <option<% if (lcase(strHoverTextDecoration)="underline" or lcase(HoverTextDecoration)="underline") then Response.Write(" selected") %>>����</option>
    </select>
    <a href="JavaScript:openWindow3('pop_config_help.asp#fontdecorations')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">���ⱳ����ɫ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><input type="text" name="strHeadCellColor" size="10" value="<% if strHeadCellColor <> "" then Response.Write(strHeadCellColor) else '## do nothing %>">
    <a href="JavaScript:openWindow3('pop_config_help.asp#colors')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">����������ɫ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><input type="text" name="strHeadFontColor" size="10" value="<% if strHeadFontColor <> "" then Response.Write(strHeadFontColor) else '## do nothing %>">
    <a href="JavaScript:openWindow3('pop_config_help.asp#colors')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">������������ɫ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><input type="text" name="strCategoryCellColor" size="10" value="<% if strCategoryCellColor <> "" then Response.Write(strCategoryCellColor) else '## do nothing %>">
    <a href="JavaScript:openWindow3('pop_config_help.asp#colors')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">������������ɫ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><input type="text" name="strCategoryFontColor" size="10" value="<% if strCategoryFontColor <> "" then Response.Write(strCategoryFontColor) else '## do nothing %>">
    <a href="JavaScript:openWindow3('pop_config_help.asp#colors')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">һ���񱳾�ɫ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><input type="text" name="strForumFirstCellColor" size="10" value="<% if strForumFirstCellColor <> "" then Response.Write(strForumFirstCellColor) else '## do nothing %>">
    <a href="JavaScript:openWindow3('pop_config_help.asp#colors')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">��һ��������ɫ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><input type="text" name="strForumCellColor" size="10" value="<% if strForumCellColor <> "" then Response.Write(strForumCellColor) else '## do nothing %>">
    <a href="JavaScript:openWindow3('pop_config_help.asp#colors')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�ڶ���������ɫ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><input type="text" name="strAltForumCellColor" size="10" value="<% if strAltForumCellColor <> "" then Response.Write(strAltForumCellColor) else '## do nothing %>">
    <a href="JavaScript:openWindow3('pop_config_help.asp#colors')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">��̳������ɫ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><input type="text" name="strForumFontColor" size="10" value="<% if strForumFontColor <> "" then Response.Write(strForumFontColor) else '## do nothing %>">
    <a href="JavaScript:openWindow3('pop_config_help.asp#colors')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">��̳����������ɫ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><input type="text" name="strForumLinkColor" size="10" value="<% if strForumLinkColor <> "" then Response.Write(strForumLinkColor) else '## do nothing %>">
    <a href="JavaScript:openWindow3('pop_config_help.asp#colors')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">���߿���ɫ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><input type="text" name="strTableBorderColor" size="10" value="<% if strTableBorderColor <> "" then Response.Write(strTableBorderColor) else '## do nothing %>">
    <a href="JavaScript:openWindow3('pop_config_help.asp#colors')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">��ҳ������ɫ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><input type="text" name="strPopUpTableColor" size="10" value="<% if strPopUpTableColor <> "" then Response.Write(strPopUpTableColor) else '## do nothing %>">
    <a href="JavaScript:openWindow3('pop_config_help.asp#colors')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">��ҳ����߿���ɫ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><input type="text" name="strPopUpBorderColor" size="10" value="<% if strPopUpBorderColor <> "" then Response.Write(strPopUpBorderColor) else '## do nothing %>">
    <a href="JavaScript:openWindow3('pop_config_help.asp#colors')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">��������ɫ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><input type="text" name="strNewFontColor" size="10" value="<% if strNewFontColor <> "" then Response.Write(strNewFontColor) else '## do nothing %>">
    <a href="JavaScript:openWindow3('pop_config_help.asp#colors')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgcolor="<% =strHeadCellColor %>" colspan="2"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>"><b>����С����</b></font></td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�������п�ȣ�</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><input type="text" name="strTopicWidthLeft" size="5" value="<% if strTopicWidthLeft <> "" then Response.Write(strTopicWidthLeft) else Response.Write("100") %>">
    <a href="JavaScript:openWindow3('pop_config_help.asp#columnwidth')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�������в��Զ�����</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    ������<input type="radio" class="radio" name="strTopicNoWrapLeft" value="1" <% if (lcase(strTopicNoWrapLeft) <> "0" or lcase(TopicNoWrapLeft) <> "0") then Response.Write("checked")%>> 
    �رգ�<input type="radio" class="radio" name="strTopicNoWrapLeft" value="0" <% if (lcase(strTopicNoWrapLeft) = "0" or lcase(TopicNoWrapLeft) = "0") then Response.Write("checked")%>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#nowrap')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </font></td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�������п�ȣ�</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><input type="text" name="strTopicWidthRight" size="5" value="<% if strTopicWidthRight <> "" then Response.Write(strTopicWidthRight) else Response.Write("100%") %>">
    <a href="JavaScript:openWindow3('pop_config_help.asp#columnwidth')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�������в��Զ����У�</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    ������<input type="radio" class="radio" name="strTopicNoWrapRight" value="1" <% if (lcase(strTopicNoWrapRight) = "1" or lcase(TopicNoWrapRight) = "1") then Response.Write("checked")%>> 
    �رգ�<input type="radio" class="radio" name="strTopicNoWrapRight" value="0" <% if (lcase(strTopicNoWrapRight) <> "1" or lcase(TopicNoWrapRight) <> "1") then Response.Write("checked")%>>
    <a href="JavaScript:openWindow3('pop_config_help.asp#nowrap')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
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
