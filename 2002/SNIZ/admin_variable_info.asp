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
<!--#INCLUDE FILE="inc_functions.asp" -->
<!--#INCLUDE file="inc_top.asp" -->

	<table border="0" align=center width="100%">
	<tr>
	    <td width="33%" align="left" nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
		<img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="default.asp">������̳��ҳ</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="admin_home.asp">��̳��������</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" height=15 width=15 border="0">&nbsp;��̳ϵͳ��Ѷ<br>
		</font></td>
	</tr>
	</table><br>
    <table border="0" cellspacing="1" cellpadding="4" align=center width="100%" bgcolor="<% =strPopUpBorderColor %>">
     <tr>
        <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>Variable Name</b></font></td>
        <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>Value</b></font></td>
      </tr>

     <tr>
        <td align="center" colspan="2" bgcolor="<% =strHeadCellColor %>"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">General information</font></b></td>
     </tr>	
     <tr>
        <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>StrCookieUrl</b></font></td>
        <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><%=ChkString(StrCookieUrl, "display")%></font></td>
      </tr>
     <tr>
        <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>StrUniqueID</b></font></td>
        <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><%=ChkString(StrUniqueID, "display")%></font></td>
      </tr>
     <tr>
        <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>strAuthType</b></font></td>
        <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><%=ChkString(strAuthType, "display")%></font></td>
      </tr>
     <tr>
        <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>strDBNTSQLName</b></font></td>
        <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><%=ChkString(strDBNTSQLName, "display")%></font></td>
      </tr>
     <tr>
        <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>STRdbntUserName</b></font></td>
        <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><%=ChkString(STRdbntUserName, "display")%></font></td>
      </tr>

     <tr>
        <td align="center" colspan="2" bgcolor="<% =strHeadCellColor %>"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">Cookies</font></b></td>
      </tr>
<% for each key in Request.Cookies 

	if Request.Cookies(key).HasKeys then
		for each subkey in Request.Cookies(key)
%>
 		     <tr>
		        <td bgColor="<% =strPopUpTableColor %>" valign="top"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b><% =ChkString(key, "display") %> (<% =ChkString(subkey, "display") %>)</b></font></td>
		        <td bgColor="<% =strPopUpTableColor %>"><font face="courier" size="<% =strDefaultFontSize %>">
<%
		if Request.Cookies(key)(subkey) = "" then
			Response.Write "&nbsp;"
		else
			Response.Write ChkString(CStr(Request.Cookies(key)(subkey)), "display")
		end if 
%>
		        </font></td>
		      </tr>
<%		next
	else
%>
 		     <tr>
		        <td bgColor="<% =strPopUpTableColor %>" valign="top"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b><% =ChkString(key, "display") %></b></font></td>
		        <td bgColor="<% =strPopUpTableColor %>"><font face="courier" size="<% =strDefaultFontSize %>">
<%
		if Request.Cookies(key) = "" then
			Response.Write "&nbsp;"
		else
			Response.Write ChkString(CStr(Request.Cookies(key)), "display")
		end if 
%>
		        </font></td>
		      </tr>
<%
	end if
next  %>
 
     <tr>
        <td align="center" colspan="2" bgcolor="<% =strHeadCellColor %>"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">Session variables</font></b></td>
      </tr>
<% for each key in Session.Contents

	if left(key, len(strCookieUrl)) = strCookieUrl or left(key, len(strUniqueID)) = strUniqueID then
%>
      <tr>
        <td bgColor="<% =strPopUpTableColor %>" valign="top"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b><% =ChkString(key, "display") %></b></font></td>
        <td bgColor="<% =strPopUpTableColor %>"><font face="courier" size="<% =strDefaultFontSize %>">
<%
	if Session.Contents(key) = "" then
		Response.Write "&nbsp;"
	else
		Response.Write ChkString(CStr(Session.Contents(key)), "display")
	end if 
%>
        </font></td>
      </tr>
<% 
	end if
next 

%>

      <tr>
        <td align="center" colspan="2" bgcolor="<% =strHeadCellColor %>"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">Application variables</font></b></td>
      </tr>
<% for each key in Application.Contents

	if left(key, len(strCookieUrl)) = strCookieUrl or left(key, len(strUniqueID)) = strUniqueID then
%>
      <tr>
        <td bgColor="<% =strPopUpTableColor %>" valign="top"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b><% = ChkString(key, "display") %></b></font></td>
        <td bgColor="<% =strPopUpTableColor %>"><font face="courier" size="<% =strDefaultFontSize %>">
<%
	if Application.Contents(key) = "" then
		Response.Write "&nbsp;"
	else
		Response.Write ChkString(CStr(Application.Contents(key)), "display")
	end if 
%>
        </font></td>
      </tr>
<% 
	end if
next 

%>
    </table>
</p>
<!--#INCLUDE file="inc_footer.asp" -->
<% Else %>
<% Response.Redirect "admin_login.asp" %>
<% End IF %>