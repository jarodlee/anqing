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
<!--#INCLUDE file="inc_functions.asp" -->
<!--#INCLUDE file="inc_top.asp" -->

<%

	Dim objDict
	Set objDict = Server.CreateObject("Scripting.Dictionary")
	set objRec = my_Conn.execute("SELECT M_CODE, M_VALUE FROM " & strTablePrefix & "MODS WHERE M_NAME = 'HModEnable'")

	while not objRec.EOF 
	objDict.Add objRec.Fields.Item("m_code").Value, objRec.Fields.Item("m_value").Value
	objRec.moveNext
	wend 

	if Request.Form("Method_Type") = "Write_Configuration" then 
	
		Err_Msg = ""
	
		if Err_Msg = "" then

			'## Forum_SQL
			objDict.Item("Attachment") = Request.Form("strAllowUploads")
			objDict.Item("Pmessages")= Request.Form("strPmessages")
			objDict.Item("UserFields")= Request.Form("strUserFields")
			objDict.Item("PollMentor")= Request.Form("strPollMentor")
			objDict.Item("SideMenu")= Request.Form("strSideMenu")
			objDict.Item("NewMemberPM")= Request.Form("strOnOff")
			objDict.Item("imageURLPath")= Request.Form("fstrImageURL")
			a = objDict.Keys
			b = objDict.Items
   			For i = 0 To objDict.Count -1 ' Iterate the array.
			strSql = "UPDATE " & strTablePrefix & "MODS "
			strSql = strSql & " SET M_VALUE ='" & b(i) & "'"
			strSQL = strSql & " WHERE M_NAME = 'HModEnable' AND M_CODE = '" & a(i) & "' "

			my_Conn.Execute (strSql)
   			Next
			Application(strCookieURL & "ConfigLoaded") = ""

%>
<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">�����Ѿ�����!</font></p>


<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">�޸���ɣ�</font></p>

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
<%		end if 
	else
%>
<form action="admin_mods.asp" method="post" id="Form1" name="Form1">
<input type="hidden" name="Method_Type" value="Write_Configuration">
<table border="0" cellspacing="1" cellpadding="3" align=center bgcolor="<% =strPopUpBorderColor %>">
  <tr valign="top">
    <td bgcolor="<% =strHeadCellColor %>" colspan="2"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>"><b>HuwR Mod Configuration</b></font></td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�ϴ��ļ�:&nbsp;</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    ������<input type="radio" class="radio" name="strAllowUploads" value="1" <% if intAllowUploads <> 0 then Response.Write("checked") %>> 
    �رգ�<input type="radio" class="radio" name="strAllowUploads" value="0" <% if intAllowUploads = 0 then Response.Write("checked") %>>

    </td>
  </tr>
    <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">���Ļ���Ϣ:&nbsp;</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    ������<input type="radio" class="radio" name="strPMessages" value="1" <% if intPMessages <> 0 then Response.Write("checked") %>> 
    �رգ�<input type="radio" class="radio" name="strPMessages" value="0" <% if intPMessages = 0 then Response.Write("checked") %>>

    </td>
  </tr>
  </tr>
    <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�����Ա������:&nbsp;</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    ������<input type="radio" class="radio" name="strUserFields" value="1" <% if intUserFields <> 0 then Response.Write("checked") %>> 
    �رգ�<input type="radio" class="radio" name="strUserFields" value="0" <% if intUserFields = 0 then Response.Write("checked") %>>

    </td>
  </tr>
  </tr>
    <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">��̳ͶƱ:&nbsp;</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    ������<input type="radio" class="radio" name="strPollMentor" value="1" <% if intPollMentor <> 0 then Response.Write("checked") %>> 
    �رգ�<input type="radio" class="radio" name="strPollMentor" value="0" <% if intPollMentor = 0 then Response.Write("checked") %>>

    </td>
  </tr>
    <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">��ݹ��ܲ˵�:&nbsp;</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    ������<input type="radio" class="radio" name="strSideMenu" value="1" <% if intSideMenu <> 0 then Response.Write("checked") %>> 
    �رգ�<input type="radio" class="radio" name="strSideMenu" value="0" <% if intSideMenu = 0 then Response.Write("checked") %>>

    </td>
  </tr>
   <tr valign="top">
       <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�������Ļ����»�Ա:&nbsp;</font></td>
		   <td bgColor="<% =strPopUpTableColor %>">
		   ������<input type="radio" class="radio" name="strOnOff" value="1" <% if intNewMemberPM <> 0 then Response.Write("checked") %>> 
		   �رգ�<input type="radio" class="radio" name="strOnOff" value="0" <% if intNewMemberPM = 0 then Response.Write("checked") %>>
	   </td>
   </tr>
  
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">ImageͼƬ·��:&nbsp;</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    <input type="text" name="fstrImageURL" value="<%=strImageURL %>" >
    </td>
  </tr>

  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" colspan="2" align="center"><input type="submit" value="�ύ����" id="submit1" name="submit1"> <input type="reset" value="��������" id="reset1" name="reset1"></td>
  </tr>
</table>
</form>
</font>
<%	end if %>
<div align="center">    <p><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="admin_home.asp">���ع�������</A></font></p></div>
    </font>
    </center>
    </td>
  </tr>
</table>

</BODY>
</HTML>
<!--#include file="inc_footer.asp"-->
