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
<!--#INCLUDE FILE="inc_code.asp" -->

<script language="JavaScript">
<!--
function OpenPreview()
{
	//var curCookie = "strSignaturePreview=" + escape(document.Form1.strMessage.value);
	var curCookie = "strMessagePreview=" + escape(document.Form1.strMessage.value);
	document.cookie = curCookie;
	popupwin = window.open('pop_preview.asp', 'preview_page', 'scrollbars=yes,width=450,height=250')	
}
//-->
</script>

<%
	Dim objDict
	Set objDict = Server.CreateObject("Scripting.Dictionary")
	set objRec = my_Conn.execute("SELECT * FROM " & strTablePrefix & "MODS WHERE M_NAME = 'NewMemPM'")

	while not objRec.EOF 
	objDict.Add objRec.Fields.Item("m_code").Value, objRec.Fields.Item("m_value").Value
	objRec.moveNext
	wend

	if Request.Form("Method_Type") = "Write_Configuration" then 
		Err_Msg = ""

		if Err_Msg = "" then
			'## Forum_SQL
			objDict.Item("Admin") = Request.Form("strAdmin")
			objDict.Item("Subject") = Request.Form("strSubject")
			objDict.Item("Message") = Request.Form("strMessage")
			a = objDict.Keys
			b = objDict.Items

   			For i = 0 To objDict.Count -1 ' Iterate the array.

			strSql = "UPDATE " & strTablePrefix & "MODS "
			strSql = strSql & " SET M_VALUE ='" & b(i) & "'"
			strSQL = strSql & " WHERE M_NAME = 'NewMemPM' AND M_CODE = '" & a(i) & "' "
			'response.write strSQL
			my_Conn.Execute (strSql)
   			Next

			Application(strCookieURL & "ConfigLoaded") = ""
%>
<p align="center">
   <font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">
   �����Ѿ�����!
   </font>
</p>

<p align="center">
   <font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">
   �޸���ɣ�
   </font>
</p>

<%		else %>

<p align=center>
   <font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">
   ������������������û����д����
   </font>
</p>

<table align=center border=0>
  <tr>
    <td>
		<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
		<ul>
		<% =Err_Msg %>
		</ul>
	    </font>
	</td>
  </tr>
</table>

<p align=center>
	<font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>>
	<a href="JavaScript:history.go(-1)">�뷵����������</a>
	</font>
</p>

<%		end if %>

<%	else 
			strSql = "SELECT * FROM " & strTablePrefix & "CONFIG_NEWMEMPM "

'Response.Write strSql
'Response.End

%>
<table border="0" width="100%">
  <tr>
    <td width="33%" align="left" nowrap>
		<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    	<img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="mainforum.asp">������̳��ҳ</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="admin_home.asp">��̳��������</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" height=15 width=15 border="0">&nbsp;���ͻ�ӭ�����»�Ա</font>
	</td>
  </tr>
</table>

<form action="admin_new_mem_pm.asp" method="post" id="Form1" name="Form1">
<input type="hidden" name="Method_Type" value="Write_Configuration">
<input type="hidden" name="cookies" value="yes">

<table border="0" cellspacing="1" cellpadding="4" align=center bgcolor="<% =strPopUpBorderColor %>">
			   <tr valign="top">
			   	   <td bgcolor="<%=strHeadCellColor%>" align="left">
				   	   <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">
			���ͻ�ӭ�����»�Ա
					   </font>
				   </td>
			   </tr>
			   <tr>
  			  	   <td bgColor="<% =strPopUpTableColor %>" noWrap vAlign="top" align="right">
				   	   <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">Ĭ�ϻ�ӭ�ʷ�����:</font>
					   <input maxLength="50" name="strAdmin" value="<% if objDict.Item("Admin") <> "" then response.write(objDict.Item("Admin")) else response.write("admin") %>" size="40">
					   <a href="JavaScript:openWindow3('pop_config_help.asp#email')">
					   <img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
				   </td>
			   </tr>
			   <tr>
			       <td bgColor="<% =strPopUpTableColor %>" noWrap vAlign="top" align="right">
				   	   <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">��ӭ�ʱ��⣺</font>
					   <input maxLength="50" name="strSubject" value="<% if objDict.Item("Subject") <> "" then response.write(objDict.Item("Subject")) else response.write("Welcome to our forum") %>" size="40">
					   <a href="JavaScript:openWindow3('pop_config_help.asp#email')">
					   <img src="<%=strImageURL %>icon_smile_question.gif" border="0"></a>
				   </td>
			   </tr>
			   <tr>
			   	   <td bgColor="<% =strPopUpTableColor %>" align="center">
				   	   <textarea cols="40" name="strMessage" rows="15" wrap="VIRTUAL"><% if objDict.Item("Message") <> "" then response.write(objDict.Item("Message")) else response.write("No Welcome letter has been written. Type what you would like the user to see when they sign up for a new membership") %></textarea>
					   <br>
				   </td>
			   </tr>
			   <tr>
			   	   <td bgColor="<% =strPopUpTableColor %>" align="center">
				   	   <input name="Preview" type="button" value=" Ԥ �� " onclick="OpenPreview()">
					   &nbsp;
					   <input name="Reset" type="reset" value="ȫ�����"></td>
				   </td>
			   </tr>
			   <tr valign="top">
			   	   <td bgColor="<% =strPopUpTableColor %>" align="center">
				   	   <input type="submit" value="�ύ����" id="submit1" name="submit1">
					   &nbsp;
					   <input type="reset" value="��������" id="reset1" name="reset1">
				   </td>
			   </tr>
		</table>
</form>

<%	end if %>

    <p align="center">
	<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
	<a href="admin_home.asp">���ع�������</A>
	</font>
	</p>

    </td>
  </tr>
</table>

<!--#INCLUDE file="inc_footer.asp" -->
