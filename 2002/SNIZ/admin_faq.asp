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
<!--#include file="inc_code.asp" -->
<!--#INCLUDE file="inc_top.asp" -->
<script language="JavaScript">
<!--
//#################################################################################
//## Allowed User - Selection Code
//#################################################################################

function selectUsers()
{
	if (document.PostTopic.AuthUsers.length == 1)
	{
		document.PostTopic.AuthUsers.options[0].value = "";
		return;
	}
	if (document.PostTopic.AuthUsers.length == 2)
		document.PostTopic.AuthUsers.options[0].selected = true
	else
	for (x = 0;x < document.PostTopic.AuthUsers.length - 1 ;x++)
		document.PostTopic.AuthUsers.options[x].selected = true;
}

function MoveWholeList(strAction)
{
	if (strAction == "Add")
	{
		if (document.PostTopic.AuthUsersCombo.length > 1)
		{
		for (x = 0;x < document.PostTopic.AuthUsersCombo.length - 1 ;x++)
			document.PostTopic.AuthUsersCombo.options[x].selected = true;
			InsertSelection("Add");
		}
	}
	else
	{
		if (document.PostTopic.AuthUsers.length > 1)
		{
		for (x = 0;x < document.PostTopic.AuthUsers.length - 1 ;x++)
			document.PostTopic.AuthUsers.options[x].selected = true;
			InsertSelection("Del");
		}
	}
}

function InsertSelection(strAction)
{
	var pos,user,mText;
	var count,finished;

	if (strAction == "Add")
	{
		pos = document.PostTopic.AuthUsers.length;
		finished = false;
		count = 0;	
		do //Add to destination
		{
			if (document.PostTopic.AuthUsersCombo.options[count].text == "")
			{
				finished = true;
				continue;
			}
			if (document.PostTopic.AuthUsersCombo.options[count].selected)
			{
				document.PostTopic.AuthUsers.length +=1;
				document.PostTopic.AuthUsers.options[pos].value = document.PostTopic.AuthUsers.options[pos-1].value;	
				document.PostTopic.AuthUsers.options[pos].text = document.PostTopic.AuthUsers.options[pos-1].text;
				document.PostTopic.AuthUsers.options[pos-1].value = document.PostTopic.AuthUsersCombo.options[count].value;	
				document.PostTopic.AuthUsers.options[pos-1].text = document.PostTopic.AuthUsersCombo.options[count].text;
				document.PostTopic.AuthUsers.options[pos-1].selected = true;
			}
			pos = document.PostTopic.AuthUsers.length;
			count += 1;
		}while (!finished); //finished adding
		finished = false;
		count = document.PostTopic.AuthUsersCombo.length - 1;
		do //remove from source
		{	
			if (document.PostTopic.AuthUsersCombo.options[count].text == "")
			{
				--count;
				continue;
			}
			if (document.PostTopic.AuthUsersCombo.options[count].selected )
			{
				for ( z = count ; z < document.PostTopic.AuthUsersCombo.length-1;z++)
				{	
					document.PostTopic.AuthUsersCombo.options[z].value = document.PostTopic.AuthUsersCombo.options[z+1].value;	
					document.PostTopic.AuthUsersCombo.options[z].text = document.PostTopic.AuthUsersCombo.options[z+1].text;
				}
				document.PostTopic.AuthUsersCombo.length -= 1;
			}
			--count;
			if (count < 0)
				finished = true;
		}while(!finished) //finished removing
	}	

	if (strAction == "Del")
	{
		pos = document.PostTopic.AuthUsersCombo.length;
		finished = false;
		count = 0;	
		do //Add to destination
		{
			if (document.PostTopic.AuthUsers.options[count].text == "")
			{
				finished = true;
				continue;
			}
			if (document.PostTopic.AuthUsers.options[count].selected)
			{
				document.PostTopic.AuthUsersCombo.length +=1;
				document.PostTopic.AuthUsersCombo.options[pos].value = document.PostTopic.AuthUsersCombo.options[pos-1].value;	
				document.PostTopic.AuthUsersCombo.options[pos].text = document.PostTopic.AuthUsersCombo.options[pos-1].text;
				document.PostTopic.AuthUsersCombo.options[pos-1].value = document.PostTopic.AuthUsers.options[count].value;	
				document.PostTopic.AuthUsersCombo.options[pos-1].text = document.PostTopic.AuthUsers.options[count].text;
				document.PostTopic.AuthUsersCombo.options[pos-1].selected = true;
			}
			count += 1;
			pos = document.PostTopic.AuthUsersCombo.length;
		}while (!finished); //finished adding
		finished = false;
		count = document.PostTopic.AuthUsers.length - 1;
		do //remove from source
		{	
			if (document.PostTopic.AuthUsers.options[count].text == "")
			{
				--count;
				continue;
			}
			if (document.PostTopic.AuthUsers.options[count].selected )
			{
				for ( z = count ; z < document.PostTopic.AuthUsers.length-1;z++)
				{	
					document.PostTopic.AuthUsers.options[z].value = document.PostTopic.AuthUsers.options[z+1].value;	
					document.PostTopic.AuthUsers.options[z].text = document.PostTopic.AuthUsers.options[z+1].text;
				}
				document.PostTopic.AuthUsers.length -= 1;
			}
			--count;
			if (count < 0)
				finished = true;
		}while(!finished) //finished removing
	}	
}

function autoReload(objform)
{
	var tmpCookieURL = '<%=strCookieURL%>';
	if (objform.SelectSize.value == 1)
	{
		document.PostTopic.Message.cols = 45;
		document.PostTopic.Message.rows = 6;
	}
	if (objform.SelectSize.value == 2)
	{
		document.PostTopic.Message.cols = 80;
		document.PostTopic.Message.rows = 12;
	}
	if (objform.SelectSize.value == 3)
	{
		document.PostTopic.Message.cols = 90;
		document.PostTopic.Message.rows = 12;
	}
	if (objform.SelectSize.value == 4)
	{
		document.PostTopic.Message.cols = 130;
		document.PostTopic.Message.rows = 15;
	}
	document.cookie = tmpCookieURL + "strSelectSize=" + objform.SelectSize.value
}

function OpenPreview()
{
	var curCookie = "strMessagePreview=" + escape(document.PostTopic.Message.value);
	document.cookie = curCookie;
	popupWin = window.open('pop_preview.asp', 'preview_page', 'scrollbars=yes,width=750,height=450')	
}
//-->
</script>
<table border="0" width="100%">
  <tr>
    <td width="33%" align="left" nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    <img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="default.asp">������̳��ҳ</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="admin_home.asp">��̳��������</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" height=15 width=15 border="0">&nbsp;�����������<br>
    </font></td>
  </tr>
</table>



<p><center><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="admin_faq.asp">�޸�ɾ��</a> | <a href="admin_faq.asp?action=submit">�����ʴ�</a></font></center><p>
<%
strfaction = request.querystring("action")

Select Case strfaction
	Case ""
    
    strsql = "SELECT F_ID, F_FAQ_QUESTION, F_FAQ_ANSWER FROM " & strTablePrefix & "FAQ ORDER BY F_ID DESC"
    set frs = my_conn.execute(strsql)
    strdisplay = ""
    response.write("<table align=""center"" border=0 cellspacing=1 cellpadding=3 bgcolor=""" & strTableBorderColor & """>" & vbcrlf)
    	if frs.eof then
        	strdisplay = strdisplay & "<tr bgcolor=""" & strForumCellColor & """><td align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>��ʱû���κ��ʴ�, <a href=""admin_faq.asp?action=submit"">���������</a></font></td></tr>"
        else
        	Do until frs.eof
           ' ChkString(Request.QueryString("FORUM_Title"),"display")
           ' response.write("<li><a href=""#faqid" & frs("F_ID") & """>" & frs("F_FAQ_QUESTION") & "</a></li>" & vbcrlf)
            strdisplay = strdisplay & "<tr><td bgcolor=""" & strCategoryCellColor & """><a href=""#top""><img src=""images/icon_go_up.gif"" align=""right"" border=""0""></a>"
			strdisplay = strdisplay & "<font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """>"
			strdisplay = strdisplay & "<b><a name=""faqid" & frs("F_ID") & """>" & ChkString(frs("F_FAQ_QUESTION"),"display") & "</a></b></font></td></tr>"
            strdisplay = strdisplay & vbcrlf
            strdisplay = strdisplay & "<tr><td bgcolor=""" & strForumCellColor & """><font size=" & strDefaultFontSize & " face=""" & strDefaultFontFace & """ color=""" & strForumFontColor & """>" & vbcrlf
			strdisplay = strdisplay & "<p>" & formatStr(frs("F_FAQ_ANSWER")) & "</p>"
            strdisplay = strdisplay & "<table width=100% border=0><tr><td align=right><form method=""post"" action=""admin_faq.asp?action=edit""><input name=""id"" type=""hidden"" value=""" & frs("F_ID") & """><input type=""submit"" value="" �޸� ""></form></td><td align=left><form method=""post"" action=""admin_faq.asp?action=delete""><input name=""id"" type=""hidden"" value=""" & frs("F_ID") & """><input type=""submit"" value="" ɾ�� ""></form></td></tr></table></td></tr>"
            frs.movenext
            Loop
        End if
    if strdisplay <> "" then
    	response.write(strdisplay)
    End if
    response.write("</table>")
    set frs = nothing

	Case "delete"
    %>
    <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    <%
    	strFID = request("id")
        if request.querystring("id") = "" then
        	response.write("<p><center>�Ƿ�Ҫɾ�������ʴ�?<br>" & _
            "<a href=""admin_faq.asp?action=delete&id=" & strFID & """>��</a> | <a href=""admin_faq.asp"">��</a></center>")
        Else
        	strsql = "DELETE FROM " & strTablePrefix & "FAQ WHERE F_ID=" & strFID
            my_conn.execute(strsql)
            response.write("<p><center>�Ѿ��ɹ�ɾ��һ���ʴ�<br><a href=""admin_faq.asp"">�������</a> ���س����������</center><p>")
        End if    	
    %>
    </font>
    <%
    Case "edit"
    	strFID = request("id")
        If request.querystring("id") = "" then
        %><center><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
            <% if strAllowForumCode = "1" then %>
            * <a href="JavaScript:openWindow3('pop_forum_code.asp')">��̳ר�ô���</a> ����<br>
			<% else %>
            * ��̳ר�ô��� �ر�<br>
			<% end if %>
            <% if strAllowHTML = "1" then %>
            * ����ʹ�� HTML ����<br>
			<% else %>
            * ��ֹʹ�� HTML ����<br>
			<% end if %>
            <form method="post" action="admin_faq.asp?action=edit&id=<%=strFID%>" name="PostTopic">
            <select name="font" onChange="thelp(this.options[this.selectedIndex].value)">
				<option value="1">����&nbsp;</option>
				<option selected value="2">��ȫ&nbsp;</option>
				<option value="0">����&nbsp;</option>
	  		  </select>  <a href="Javascript:bold();"><img src="images/icon_editor_bold.gif" width="22" height="22" alt="Bold" border="0"></a><a href="Javascript:italicize();"><img src="images/icon_editor_italicize.gif" width="23" height="22" alt="Italicized" border="0"></a><a href="Javascript:underline();"><img src="images/icon_editor_underline.gif" width="23" height="22" alt="Underline" border="0"></a>
<a href="Javascript:center();"><img src="images/icon_editor_center.gif" width="22" height="22" alt="Centered" border="0"></a>
<a href="Javascript:hyperlink();"><img src="images/icon_editor_url.gif" width="22" height="22" alt="Insert Hyperlink" border="0"></a><a href="Javascript:email();"><img src="images/icon_editor_email.gif" width="23" height="22" alt="Insert Email" border="0"></a><a href="Javascript:image();"><img src="images/icon_editor_image.gif" width="23" height="22" alt="Insert Image" border="0"></a>
<a href="Javascript:showcode();"><img src="images/icon_editor_code.gif" width="22" height="22" alt="Insert Code" border="0"></a><a href="Javascript:quote();"><img src="images/icon_editor_quote.gif" width="23" height="22" alt="Insert Quote" border="0"></a><a href="Javascript:list();"><img src="images/icon_editor_list.gif" width="23" height="22" alt="Insert List" border="0"></a>
<%		if lcase(strIcons) = "1" then %>
<a href="JavaScript:openWindow2('pop_icon_legend.asp')"><img src="images/icon_editor_smilie.gif" width="22" height="22" alt="Insert Smilie" border="0"></a>
<% end if %><br></center>
            <%
            strsql = "SELECT F_ID, F_FAQ_QUESTION, F_FAQ_ANSWER FROM " & strTablePrefix & "FAQ WHERE F_ID=" & strFID
    		set frs = my_conn.execute(strsql)
            
        	response.write("<table align=""center"" border=""0"">" & vbcrlf & _
            "<input type=""hidden"" name=""id"" value=""" & strFID & """>" & _
            "<tr><td><strong><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>����:</strong></td><td><input type=""text"" name=""fquestion"" maxlength=""100"" value=""" & ChkString(frs("F_FAQ_QUESTION"),"display") & """></font></td></tr>" & _
            "<tr><td valign=""top""><strong><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>��:</strong></font></td><td valign=""top"">" & _
            "<textarea name=""Message"" cols=""35"" rows=""7"">" & formatStr(frs("F_FAQ_ANSWER")) & "</textarea>" &_
            "</td></tr><tr><td colspan=2 align=""right""><input type=""submit"" value="" �޸� ""> <input type=""reset"" value="" ��д "">" & _
			"</form></table>")

        ELSE
        	strFormQ = request.form("fquestion")
            strFormA = request.form("Message")
			
            strFormQ = ChkString(strFormQ,"title")
            strFormA = ChkString(strFormA,"message")
            
            strsql = "UPDATE " & strTablePrefix & "FAQ set F_FAQ_QUESTION='" & strFormQ & "', F_FAQ_ANSWER='" & strFormA & "' WHERE F_ID=" & strFID
			my_conn.execute(strsql)
            response.write("<p><center>�ʴ��Ѿ��޸����, <a href=""admin_faq.asp"">�������</a> ���س����������</center><p>")
        END IF
        %>
        </font>
        <%
    Case "submit"
    	IF request.querystring("do") = "" then
        	%><center>
            <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
            <% if strAllowForumCode = "1" then %>
            * <a href="JavaScript:openWindow3('pop_forum_code.asp')">��̳ר�ô���</a> ����<br>
			<% else %>
            * ��̳ר�ô��� �ر�<br>
			<% end if %>
            <% if strAllowHTML = "1" then %>
            * ����ʹ�� HTML ����<br>
			<% else %>
            * ��ֹʹ�� HTML ����<br>
			<% end if %>
            <form method="post" action="admin_faq.asp?action=submit&do=yes" name="PostTopic">
            <select name="font" onChange="thelp(this.options[this.selectedIndex].value)">
				<option value="1">����&nbsp;</option>
				<option value="2">��ȫ&nbsp;</option>
				<option selected value="0">����&nbsp;</option>
	  		  </select>  <a href="Javascript:bold();"><img src="images/icon_editor_bold.gif" width="22" height="22" alt="Bold" border="0"></a><a href="Javascript:italicize();"><img src="images/icon_editor_italicize.gif" width="23" height="22" alt="Italicized" border="0"></a><a href="Javascript:underline();"><img src="images/icon_editor_underline.gif" width="23" height="22" alt="Underline" border="0"></a>
<a href="Javascript:center();"><img src="images/icon_editor_center.gif" width="22" height="22" alt="Centered" border="0"></a>
<a href="Javascript:hyperlink();"><img src="images/icon_editor_url.gif" width="22" height="22" alt="Insert Hyperlink" border="0"></a><a href="Javascript:email();"><img src="images/icon_editor_email.gif" width="23" height="22" alt="Insert Email" border="0"></a><a href="Javascript:image();"><img src="images/icon_editor_image.gif" width="23" height="22" alt="Insert Image" border="0"></a>
<a href="Javascript:showcode();"><img src="images/icon_editor_code.gif" width="22" height="22" alt="Insert Code" border="0"></a><a href="Javascript:quote();"><img src="images/icon_editor_quote.gif" width="23" height="22" alt="Insert Quote" border="0"></a><a href="Javascript:list();"><img src="images/icon_editor_list.gif" width="23" height="22" alt="Insert List" border="0"></a>
<%		if lcase(strIcons) = "1" then %>
<a href="JavaScript:openWindow2('pop_icon_legend.asp')"><img src="images/icon_editor_smilie.gif" width="22" height="22" alt="Insert Smilie" border="0"></a>
<% end if %><br></center>
            <%
        	response.write("<table align=""center"" border=""0"">" & vbcrlf & _
            "<tr><td><strong><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>����:</strong></td><td><input type=""text"" name=""fquestion"" maxlength=""100""></font></td></tr>" & _
            "<tr><td valign=""top""><strong><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>��:</strong></font></td><td valign=""top"">" & vbcrlf &  _
            "<textarea name=""Message"" cols=""35"" rows=""7""></textarea>" &vbcrlf & _
            "</td></tr><tr><td colspan=2 align=""right""><input type=""submit"" value="" �ύ ""> <input type=""reset"" value="" ��д "">" & vbcrlf &  _
			"</table></form>")        
        
        ELSE
        	strFormQ = request.form("fquestion")
            strFormA = request.form("Message")
			
            strFormQ = ChkString(strFormQ,"title")
            strFormA = ChkString(strFormA,"message")
            
            strsql = "INSERT INTO " & strTablePrefix & "FAQ (F_FAQ_QUESTION, F_FAQ_ANSWER)"
            strsql = strsql & " VALUES ('" & strFormQ & "', '" & strFormA & "')"
            my_conn.execute(strsql)        
            response.write("<p><center>������, <a href=""admin_faq.asp"">�������</a> ���س����������</center><p>")        
        end if
        %>
        </font>
        <%
End Select
%>
<!--#INCLUDE file="inc_footer.asp" -->
<% Else %>
<% Response.Redirect "admin_login.asp" %>
<% End IF %>
