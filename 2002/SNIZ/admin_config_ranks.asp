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
    <img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="default.asp">������̳��ҳ</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="admin_home.asp">��̳��������</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" height=15 width=15 border="0">&nbsp;�ȼ��趨<br>
    </font></td>
  </tr>
</table>
<%
	if Request.Form("Method_Type") = "Write_Configuration" then 
		Err_Msg = ""
		if Request.Form("strRankAdmin") = "" then 
			Err_Msg = Err_Msg & "<li>������������Ա�ƺ�</li>"
		end if
		if Request.Form("strRankMod") = "" then 
			Err_Msg = Err_Msg & "<li>�������������ƺ�</li>"
		end if
		if Request.Form("strRankLevel0") = "" then 
			Err_Msg = Err_Msg & "<li>�����������ע���Ա�ƺ�</li>"
		end if
		if Request.Form("strRankLevel1") = "" then 
			Err_Msg = Err_Msg & "<li>���������ȼ�һ��Ա�ƺ�</li>"
		end if
		if Request.Form("strRankLevel2") = "" then 
			Err_Msg = Err_Msg & "<li>���������ȼ�����Ա�ƺ�</li>"
		end if
		if Request.Form("strRankLevel3") = "" then 
			Err_Msg = Err_Msg & "<li>���������ȼ�����Ա�ƺ�</li>"
		end if
		if Request.Form("strRankLevel4") = "" then 
			Err_Msg = Err_Msg & "<li>���������ȼ��Ļ�Ա�ƺ�</li>"
		end if
		if Request.Form("strRankLevel5") = "" then 
			Err_Msg = Err_Msg & "<li>���������ȼ����Ա�ƺ�</li>"
		end if
		if cint(Request.Form("intRankLevel1")) > cint(Request.Form("intRankLevel2")) then 
			Err_Msg = Err_Msg & "<li>�ȼ�һ���ܱȵȼ�����</li>"
		end if
		if cint(Request.Form("intRankLevel1")) > cint(Request.Form("intRankLevel3")) then 
			Err_Msg = Err_Msg & "<li>�ȼ�һ���ܱȵȼ�����</li>"
		end if
		if cint(Request.Form("intRankLevel2")) > cint(Request.Form("intRankLevel3")) then 
			Err_Msg = Err_Msg & "<li>�ȼ������ܱȵȼ�����</li>"
		end if
		if cint(Request.Form("intRankLevel1")) > cint(Request.Form("intRankLevel4")) then 
			Err_Msg = Err_Msg & "<li>�ȼ�һ���ܱȵȼ��ĸ�</li>"
		end if
		if cint(Request.Form("intRankLevel2")) > cint(Request.Form("intRankLevel4")) then 
			Err_Msg = Err_Msg & "<li>�ȼ������ܱȵȼ��ĸ�</li>"
		end if
		if cint(Request.Form("intRankLevel3")) > cint(Request.Form("intRankLevel4")) then 
			Err_Msg = Err_Msg & "<li>�ȼ������ܱȵȼ��ĸ�</li>"
		end if
		if cint(Request.Form("intRankLevel1")) > cint(Request.Form("intRankLevel5")) then 
			Err_Msg = Err_Msg & "<li>�ȼ�һ���ܱȵȼ����</li>"
		end if
		if cint(Request.Form("intRankLevel2")) > cint(Request.Form("intRankLevel5")) then 
			Err_Msg = Err_Msg & "<li>�ȼ������ܱȵȼ����</li>"
		end if
		if cint(Request.Form("intRankLevel3")) > cint(Request.Form("intRankLevel5")) then 
			Err_Msg = Err_Msg & "<li>�ȼ������ܱȵȼ����</li>"
		end if
		if cint(Request.Form("intRankLevel4")) > cint(Request.Form("intRankLevel5")) then 
			Err_Msg = Err_Msg & "<li>�ȼ��Ĳ��ܱȵȼ����</li>"
		end if

		if Err_Msg = "" then

			'## Forum_SQL
			strSql = "UPDATE " & strTablePrefix & "CONFIG "
			strSql = strSql & " SET C_STRSHOWRANK = " & Request.Form("strShowRank") & ""
			strSql = strSql & ",    C_STRRANKADMIN = '" & ChkString(Request.Form("strRankAdmin"),"name") & "'"
			strSql = strSql & ",    C_STRRANKMOD = '" & ChkString(Request.Form("strRankMod"),"name") & "'"
			strSql = strSql & ",    C_STRRANKLEVEL0 = '" & ChkString(Request.Form("strRankLevel0"),"name") & "'"
			strSql = strSql & ",    C_STRRANKLEVEL1 = '" & ChkString(Request.Form("strRankLevel1"),"name") & "'"
			strSql = strSql & ",    C_STRRANKLEVEL2 = '" & ChkString(Request.Form("strRankLevel2"),"name") & "'"
			strSql = strSql & ",    C_STRRANKLEVEL3 = '" & ChkString(Request.Form("strRankLevel3"),"name") & "'"
			strSql = strSql & ",    C_STRRANKLEVEL4 = '" & ChkString(Request.Form("strRankLevel4"),"name") & "'"
			strSql = strSql & ",    C_STRRANKLEVEL5 = '" & ChkString(Request.Form("strRankLevel5"),"name") & "'"
			strSql = strSql & ",    C_STRRANKCOLORADMIN = '" & ChkString(Request.Form("strRankColorAdmin"),"name") & "'"
			strSql = strSql & ",    C_STRRANKCOLORMOD = '" & ChkString(Request.Form("strRankColorMod"),"name") & "'"
			strSql = strSql & ",    C_STRRANKCOLOR0 = '" & ChkString(Request.Form("strRankColor0"),"name") & "'"
			strSql = strSql & ",    C_STRRANKCOLOR1 = '" & ChkString(Request.Form("strRankColor1"),"name") & "'"
			strSql = strSql & ",    C_STRRANKCOLOR2 = '" & ChkString(Request.Form("strRankColor2"),"name") & "'"
			strSql = strSql & ",    C_STRRANKCOLOR3 = '" & ChkString(Request.Form("strRankColor3"),"name") & "'"
			strSql = strSql & ",    C_STRRANKCOLOR4 = '" & ChkString(Request.Form("strRankColor4"),"name") & "'"
			strSql = strSql & ",    C_STRRANKCOLOR5 = '" & ChkString(Request.Form("strRankColor5"),"name") & "'"
			strSql = strSql & ",    C_INTRANKLEVEL0 = " & ChkString(Request.Form("intRankLevel0"),"number") & ""
			strSql = strSql & ",    C_INTRANKLEVEL1 = " & ChkString(Request.Form("intRankLevel1"),"number") & ""
			strSql = strSql & ",    C_INTRANKLEVEL2 = " & ChkString(Request.Form("intRankLevel2"),"number") & ""
			strSql = strSql & ",    C_INTRANKLEVEL3 = " & ChkString(Request.Form("intRankLevel3"),"number") & ""
			strSql = strSql & ",    C_INTRANKLEVEL4 = " & ChkString(Request.Form("intRankLevel4"),"number") & ""
			strSql = strSql & ",    C_INTRANKLEVEL5 = " & ChkString(Request.Form("intRankLevel5"),"number") & ""
			strSql = strSql & " WHERE CONFIG_ID = " & 1
		
			my_Conn.Execute (strSql)
			Application(strCookieURL & "ConfigLoaded") = ""

%>
<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">�ȼ��趨
�Ѹ�����ϣ�</font></p>
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
<form action="admin_config_ranks.asp" method="post" id="Form1" name="Form1">
<input type="hidden" name="Method_Type" value="Write_Configuration">
<table border="0" cellspacing="1" cellpadding="4" align=center bgcolor="<% =strPopUpBorderColor %>">
  <tr valign="top">
    <td bgcolor="<% =strHeadCellColor %>" colspan="2"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>"><b>�ȼ��趨</b></font></td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�ȼ���ʾģʽ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    <select name="strShowRank">
      <option value="0"<% if strShowRank = "0" then Response.Write(" selected")%>>��</option>
      <option value="1"<% if strShowRank = "1" then Response.Write(" selected")%>>ֻ��ʾ�ȼ�</option>
      <option value="2"<% if strShowRank = "2" then Response.Write(" selected")%>>ֻ��ʾ�Ǻ�</option>
      <option value="3"<% if strShowRank = "3" then Response.Write(" selected")%>>��ʾ�Ǻź͵ȼ�</option>
    </select>
    <a href="JavaScript:openWindow3('pop_config_help.asp#ShowRank')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">����Ա���ƣ�</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><input type="text" name="strRankAdmin" size="30" value="<% if strRankAdmin <> " " then Response.Write(strRankAdmin) end if %>">
    <img src="<%=strImageURL %>icon_smile_question.gif" border="0" alt="(Administrator)"></td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align=right><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�Ǻ�ѡ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    <input type=radio name=strRankColorAdmin value=gold<% if strRankColorAdmin = "gold" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_gold.gif border=0>
    <input type=radio name=strRankColorAdmin value=silver<% if strRankColorAdmin = "silver" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_silver.gif border=0>
    <input type=radio name=strRankColorAdmin value=bronze<% if strRankColorAdmin = "bronze" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_bronze.gif border=0>
    <input type=radio name=strRankColorAdmin value=orange<% if strRankColorAdmin = "orange" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_orange.gif border=0>
    <input type=radio name=strRankColorAdmin value=red<% if strRankColorAdmin = "red" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_red.gif border=0>
    <input type=radio name=strRankColorAdmin value=purple<% if strRankColorAdmin = "purple" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_purple.gif border=0>
    <input type=radio name=strRankColorAdmin value=blue<% if strRankColorAdmin = "blue" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_blue.gif border=0>
    <input type=radio name=strRankColorAdmin value=cyan<% if strRankColorAdmin = "cyan" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_cyan.gif border=0>
    <input type=radio name=strRankColorAdmin value=green<% if strRankColorAdmin = "green" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_green.gif border=0>
    <a href="JavaScript:openWindow3('pop_config_help.asp#RankColor')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�������ƣ�</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><input type="text" name="strRankMod" size="30" value="<% if strRankMod <> " " then Response.Write(strRankMod) end if %>">
    <img src="<%=strImageURL %>icon_smile_question.gif" border="0" alt="(Moderator)">
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align=right><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�Ǻ�ѡ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    <input type=radio name=strRankColorMod value=gold<% if strRankColorMod = "gold" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_gold.gif border=0>
    <input type=radio name=strRankColorMod value=silver<% if strRankColorMod = "silver" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_silver.gif border=0>
    <input type=radio name=strRankColorMod value=bronze<% if strRankColorMod = "bronze" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_bronze.gif border=0>
    <input type=radio name=strRankColorMod value=orange<% if strRankColorMod = "orange" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_orange.gif border=0>
    <input type=radio name=strRankColorMod value=red<% if strRankColorMod = "red" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_red.gif border=0>
    <input type=radio name=strRankColorMod value=purple<% if strRankColorMod = "purple" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_purple.gif border=0>
    <input type=radio name=strRankColorMod value=blue<% if strRankColorMod = "blue" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_blue.gif border=0>
    <input type=radio name=strRankColorMod value=cyan<% if strRankColorMod = "cyan" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_cyan.gif border=0>
    <input type=radio name=strRankColorMod value=green<% if strRankColorMod = "green" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_green.gif border=0>
    <a href="JavaScript:openWindow3('pop_config_help.asp#RankColor')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">��ע���û��ƺţ�&nbsp;</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><input type="text" name="strRankLevel0" size="30" value="<% if strRankLevel0 <> " " then Response.Write(strRankLevel0) else Response.Write("Starting Member") end if %>">
    <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>��������</b></font><input type="text" name="intRankLevel0" size="5" value="0" readonly>
    <img src="<%=strImageURL %>icon_smile_question.gif" border="0" alt="����������춵ȼ�һ�Ļ�Ա��">
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align=right><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�Ǻ�ѡ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    <input type=radio name=strRankColor0 value=gold<% if strRankColor0 = "gold" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_gold.gif border=0>
    <input type=radio name=strRankColor0 value=silver<% if strRankColor0 = "silver" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_silver.gif border=0>
    <input type=radio name=strRankColor0 value=bronze<% if strRankColor0 = "bronze" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_bronze.gif border=0>
    <input type=radio name=strRankColor0 value=orange<% if strRankColor0 = "orange" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_orange.gif border=0>
    <input type=radio name=strRankColor0 value=red<% if strRankColor0 = "red" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_red.gif border=0>
    <input type=radio name=strRankColor0 value=purple<% if strRankColor0 = "purple" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_purple.gif border=0>
    <input type=radio name=strRankColor0 value=blue<% if strRankColor0 = "blue" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_blue.gif border=0>
    <input type=radio name=strRankColor0 value=cyan<% if strRankColor0 = "cyan" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_cyan.gif border=0>
    <input type=radio name=strRankColor0 value=green<% if strRankColor0 = "green" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_green.gif border=0>
    <a href="JavaScript:openWindow3('pop_config_help.asp#RankColor')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�ȼ�һ��Ա���ƣ�</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><input type="text" name="strRankLevel1" size="30" value="<% if strRankLevel1 <> " " then Response.Write(strRankLevel1) end if %>">
    <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>��������</b></font><input type="text" name="intRankLevel1" size="5" value="<% =intRankLevel1 %>">
    <img src="<%=strImageURL %>icon_smile_question.gif" border="0" alt="����������춵ȼ�����Ա��">
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align=right><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�Ǻ�ѡ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    <input type=radio name=strRankColor1 value=gold<% if strRankColor1 = "gold" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_gold.gif border=0>
    <input type=radio name=strRankColor1 value=silver<% if strRankColor1 = "silver" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_silver.gif border=0>
    <input type=radio name=strRankColor1 value=bronze<% if strRankColor1 = "bronze" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_bronze.gif border=0>
    <input type=radio name=strRankColor1 value=orange<% if strRankColor1 = "orange" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_orange.gif border=0>
    <input type=radio name=strRankColor1 value=red<% if strRankColor1 = "red" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_red.gif border=0>
    <input type=radio name=strRankColor1 value=purple<% if strRankColor1 = "purple" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_purple.gif border=0>
    <input type=radio name=strRankColor1 value=blue<% if strRankColor1 = "blue" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_blue.gif border=0>
    <input type=radio name=strRankColor1 value=cyan<% if strRankColor1 = "cyan" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_cyan.gif border=0>
    <input type=radio name=strRankColor1 value=green<% if strRankColor1 = "green" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_green.gif border=0>
    <a href="JavaScript:openWindow3('pop_config_help.asp#RankColor')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
   </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�ȼ�����Ա���ƣ�</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><input type="text" name="strRankLevel2" size="30" value="<% if strRankLevel2 <> " " then Response.Write(strRankLevel2) end if %>">
    <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>��������</b></font><input type="text" name="intRankLevel2" size="5" value="<% =intRankLevel2 %>">
    <img src="<%=strImageURL %>icon_smile_question.gif" border="0" alt="����������춵ȼ�����Ա��">
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align=right><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�Ǻ�ѡ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    <input type=radio name=strRankColor2 value=gold<% if strRankColor2 = "gold" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_gold.gif border=0>
    <input type=radio name=strRankColor2 value=silver<% if strRankColor2 = "silver" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_silver.gif border=0>
    <input type=radio name=strRankColor2 value=bronze<% if strRankColor2 = "bronze" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_bronze.gif border=0>
    <input type=radio name=strRankColor2 value=orange<% if strRankColor2 = "orange" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_orange.gif border=0>
    <input type=radio name=strRankColor2 value=red<% if strRankColor2 = "red" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_red.gif border=0>
    <input type=radio name=strRankColor2 value=purple<% if strRankColor2 = "purple" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_purple.gif border=0>
    <input type=radio name=strRankColor2 value=blue<% if strRankColor2 = "blue" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_blue.gif border=0>
    <input type=radio name=strRankColor2 value=cyan<% if strRankColor2 = "cyan" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_cyan.gif border=0>
    <input type=radio name=strRankColor2 value=green<% if strRankColor2 = "green" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_green.gif border=0>
    <a href="JavaScript:openWindow3('pop_config_help.asp#RankColor')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�ȼ�����Ա���ƣ�</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><input type="text" name="strRankLevel3" size="30" value="<% if strRankLevel3 <> " " then Response.Write(strRankLevel3) end if %>">
    <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>��������</b></font><input type="text" name="intRankLevel3" size="5" value="<% =intRankLevel3 %>">
    <img src="<%=strImageURL %>icon_smile_question.gif" border="0" alt="����������춵ȼ��Ļ�Ա��">
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align=right><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�Ǻ�ѡ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    <input type=radio name=strRankColor3 value=gold<% if strRankColor3 = "gold" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_gold.gif border=0>
    <input type=radio name=strRankColor3 value=silver<% if strRankColor3 = "silver" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_silver.gif border=0>
    <input type=radio name=strRankColor3 value=bronze<% if strRankColor3 = "bronze" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_bronze.gif border=0>
    <input type=radio name=strRankColor3 value=orange<% if strRankColor3 = "orange" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_orange.gif border=0>
    <input type=radio name=strRankColor3 value=red<% if strRankColor3 = "red" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_red.gif border=0>
    <input type=radio name=strRankColor3 value=purple<% if strRankColor3 = "purple" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_purple.gif border=0>
    <input type=radio name=strRankColor3 value=blue<% if strRankColor3 = "blue" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_blue.gif border=0>
    <input type=radio name=strRankColor3 value=cyan<% if strRankColor3 = "cyan" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_cyan.gif border=0>
    <input type=radio name=strRankColor3 value=green<% if strRankColor3 = "green" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_green.gif border=0>
    <a href="JavaScript:openWindow3('pop_config_help.asp#RankColor')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�ȼ��Ļ�Ա���ƣ�</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><input type="text" name="strRankLevel4" size="30" value="<% if strRankLevel4 <> " " then Response.Write(strRankLevel4) end if %>">
    <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>��������</b></font><input type="text" name="intRankLevel4" size="5" value="<% =intRankLevel4 %>">
    <img src="<%=strImageURL %>icon_smile_question.gif" border="0" alt="����������춵ȼ����Ա��">
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align=right><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�Ǻ�ѡ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    <input type=radio name=strRankColor4 value=gold<% if strRankColor4 = "gold" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_gold.gif border=0>
    <input type=radio name=strRankColor4 value=silver<% if strRankColor4 = "silver" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_silver.gif border=0>
    <input type=radio name=strRankColor4 value=bronze<% if strRankColor4 = "bronze" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_bronze.gif border=0>
    <input type=radio name=strRankColor4 value=orange<% if strRankColor4 = "orange" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_orange.gif border=0>
    <input type=radio name=strRankColor4 value=red<% if strRankColor4 = "red" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_red.gif border=0>
    <input type=radio name=strRankColor4 value=purple<% if strRankColor4 = "purple" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_purple.gif border=0>
    <input type=radio name=strRankColor4 value=blue<% if strRankColor4 = "blue" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_blue.gif border=0>
    <input type=radio name=strRankColor4 value=cyan<% if strRankColor4 = "cyan" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_cyan.gif border=0>
    <input type=radio name=strRankColor4 value=green<% if strRankColor4 = "green" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_green.gif border=0>
    <a href="JavaScript:openWindow3('pop_config_help.asp#RankColor')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�ȼ����Ա���ƣ�</font></td>
    <td bgColor="<% =strPopUpTableColor %>"><input type="text" name="strRankLevel5" size="30" value="<% if strRankLevel5 <> " " then Response.Write(strRankLevel5) end if %>">
    <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>��������</b></font><input type="text" name="intRankLevel5" size="5" value="<% =intRankLevel5 %>">
    <img src="<%=strImageURL %>icon_smile_question.gif" border="0" alt="���������Ѿ��ൽ���ˣ�">
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align=right><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�Ǻ�ѡ��</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    <input type=radio name=strRankColor5 value=gold<% if strRankColor5 = "gold" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_gold.gif border=0>
    <input type=radio name=strRankColor5 value=silver<% if strRankColor5 = "silver" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_silver.gif border=0>
    <input type=radio name=strRankColor5 value=bronze<% if strRankColor5 = "bronze" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_bronze.gif border=0>
    <input type=radio name=strRankColor5 value=orange<% if strRankColor5 = "orange" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_orange.gif border=0>
    <input type=radio name=strRankColor5 value=red<% if strRankColor5 = "red" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_red.gif border=0>
    <input type=radio name=strRankColor5 value=purple<% if strRankColor5 = "purple" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_purple.gif border=0>
    <input type=radio name=strRankColor5 value=blue<% if strRankColor5 = "blue" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_blue.gif border=0>
    <input type=radio name=strRankColor5 value=cyan<% if strRankColor5 = "cyan" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_cyan.gif border=0>
    <input type=radio name=strRankColor5 value=green<% if strRankColor5 = "green" then Response.Write(" checked") %>><img src=<%=strImageURL %>icon_star_green.gif border=0>
    <a href="JavaScript:openWindow3('pop_config_help.asp#RankColor')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" colspan="2" align="center"><input type="submit" value="�ύ����" id="submit1" name="submit1"> <input type="reset" value="��������" id="reset1" name="reset1">
    </td>
  </tr>
</table></form>
</font>
<%	end if %>
<!--#INCLUDE file="inc_footer.asp" -->
<% Else %>
<% Response.Redirect "admin_login.asp" %>
<% End IF %>
