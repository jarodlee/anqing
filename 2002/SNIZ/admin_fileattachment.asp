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
	set objRec = my_Conn.execute("SELECT * FROM " & strTablePrefix & "MODS WHERE (M_NAME = 'Attachment') OR (M_CODE = 'Attachment')")

	while not objRec.EOF 
	objDict.Add objRec.Fields.Item("m_code").Value, objRec.Fields.Item("m_value").Value
	objRec.moveNext
	wend 

	if Request.Form("Method_Type") = "Write_Configuration" then 
	if Request.Form("strAllowUploads") = "" then
		intAllowUploads = 0
	else
		intAllowUploads = Request.Form("strAllowUploads")
	end if
	
		Err_Msg = ""
	
		if Err_Msg = "" then

			'## Forum_SQL
			objDict.Item("faMaxSize") = Request.Form("intMaxFileSize")
			objDict.Item("faExtensions") = Request.Form("strFileType")
			objDict.Item("Attachment") = intAllowUploads
			a = objDict.Keys
			b = objDict.Items
   			For i = 0 To objDict.Count -1 ' Iterate the array.
			strSql = "UPDATE " & strTablePrefix & "MODS "
			strSql = strSql & " SET M_VALUE ='" & b(i) & "'"
			strSQL = strSql & " WHERE M_CODE = '" & a(i) & "' "

			my_Conn.Execute (strSql)
   			Next


			Application(strCookieURL & "ConfigLoaded") = ""

%>
<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">�ϴ������Ѿ�������ϣ�</font></p>


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
<form action="admin_fileattachment.asp" method="post" id="Form1" name="Form1">
<input type="hidden" name="Method_Type" value="Write_Configuration">
<table border="0" cellspacing="1" cellpadding="3" align=center bgcolor="<% =strPopUpBorderColor %>">
  <tr valign="top">
    <td bgcolor="<% =strHeadCellColor %>" colspan="2"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">�ļ��ϴ�����</font></td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�ļ���С����:&nbsp;</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    <input type="text" name="intMaxFileSize" size="6" value="<%= cLng(objDict.Item("faMaxSize")) %>"> (Kb)

    <a href="JavaScript:openWindow3('pop_config_help.asp#FullName')"><img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�����ϴ��ļ�����:&nbsp;</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    <input type="text" name="strFileType" value="<% if objDict.Item("faExtensions") <> "" then response.write(objDict.Item("faExtensions")) else response.write(".zip|.lha") %>"> 

    </td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">�ļ��ϴ�:&nbsp;</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    ������<input type="radio" class="radio" name="strAllowUploads" value="1" <% if objDict.Item("Attachment") <> 0 then Response.Write("checked") %>> 
    �رգ�<input type="radio" class="radio" name="strAllowUploads" value="0" <% if objDict.Item("Attachment") = 0 then Response.Write("checked") %>>

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
<div id="divMenu" visibility:hidden; background-color:F0F0F0"></div>

</BODY>
</HTML>
<!--#include file="inc_footer.asp"-->