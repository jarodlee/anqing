<%
'########## Snitz Forums 2000 Version 3.1 SR4 ####################
'#                                                               #
'#  汉化修改: 资源搜罗站                                         #
'#  电子邮件: cgier@21cn.com                                     #
'#  主页地址: http://www.sdsea.com                               #
'#            http://www.99ss.net                                #
'#            http://www.cdown.net                               #
'#	     http://www.wzdown.com                               #
'#	     http://www.13888.net                                #
'#  论坛地址:http://ubb.yesky.net                                #
'#  最后修改日期: 2001/03/12    中文版本：Version 3.1 SR4        #
'#################################################################
'# 原始来源                                                      #
'# Snitz Forums 2000 Version 3.1 SR4                             #
'# Copyright 2000 http://forum.snitz.com - All Rights Reserved   #
'#################################################################
'#【版权声明】                                                   #
'#                                                               #
'# 本软体为共享软体(shareware)提供个人网站免费使用，请勿非法修改,#
'# 转载，散播，或用于其他图利行为，并请勿删除版权声明。          #
'# 如果您的网站正式起用了这个脚本，请您通知我们，以便我们能够知晓#
'# 如果可能，请在您的网站做上我们的链接，希望能给予合作。谢谢！  #
'#################################################################
'# 请您尊重我们的劳动和版权，不要删除以上的版权声明部分，谢谢合作#
'# 如有任何问题请到我们的论坛告诉我们                            #
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
    <img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="default.asp">返回论坛首页</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open.gif" height=15 width=15 border="0">&nbsp;<a href="admin_home.asp">论坛管理中心</a>&nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" height=15 width=15 border="0">&nbsp;常见问题管理<br>
    </font></td>
  </tr>
</table>



<p><center><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="admin_faq.asp">修改删除</a> | <a href="admin_faq.asp?action=submit">新增问答</a></font></center><p>
<%
strfaction = request.querystring("action")

Select Case strfaction
	Case ""
    
    strsql = "SELECT F_ID, F_FAQ_QUESTION, F_FAQ_ANSWER FROM " & strTablePrefix & "FAQ ORDER BY F_ID DESC"
    set frs = my_conn.execute(strsql)
    strdisplay = ""
    response.write("<table align=""center"" border=0 cellspacing=1 cellpadding=3 bgcolor=""" & strTableBorderColor & """>" & vbcrlf)
    	if frs.eof then
        	strdisplay = strdisplay & "<tr bgcolor=""" & strForumCellColor & """><td align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>暂时没有任何问答, <a href=""admin_faq.asp?action=submit"">按这里添加</a></font></td></tr>"
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
            strdisplay = strdisplay & "<table width=100% border=0><tr><td align=right><form method=""post"" action=""admin_faq.asp?action=edit""><input name=""id"" type=""hidden"" value=""" & frs("F_ID") & """><input type=""submit"" value="" 修改 ""></form></td><td align=left><form method=""post"" action=""admin_faq.asp?action=delete""><input name=""id"" type=""hidden"" value=""" & frs("F_ID") & """><input type=""submit"" value="" 删除 ""></form></td></tr></table></td></tr>"
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
        	response.write("<p><center>是否要删除这条问答?<br>" & _
            "<a href=""admin_faq.asp?action=delete&id=" & strFID & """>是</a> | <a href=""admin_faq.asp"">否</a></center>")
        Else
        	strsql = "DELETE FROM " & strTablePrefix & "FAQ WHERE F_ID=" & strFID
            my_conn.execute(strsql)
            response.write("<p><center>已经成功删除一条问答<br><a href=""admin_faq.asp"">点击这里</a> 返回常见问题管理</center><p>")
        End if    	
    %>
    </font>
    <%
    Case "edit"
    	strFID = request("id")
        If request.querystring("id") = "" then
        %><center><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
            <% if strAllowForumCode = "1" then %>
            * <a href="JavaScript:openWindow3('pop_forum_code.asp')">论坛专用代码</a> 开启<br>
			<% else %>
            * 论坛专用代码 关闭<br>
			<% end if %>
            <% if strAllowHTML = "1" then %>
            * 允许使用 HTML 代码<br>
			<% else %>
            * 禁止使用 HTML 代码<br>
			<% end if %>
            <form method="post" action="admin_faq.asp?action=edit&id=<%=strFID%>" name="PostTopic">
            <select name="font" onChange="thelp(this.options[this.selectedIndex].value)">
				<option value="1">帮助&nbsp;</option>
				<option selected value="2">完全&nbsp;</option>
				<option value="0">基本&nbsp;</option>
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
            "<tr><td><strong><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>问题:</strong></td><td><input type=""text"" name=""fquestion"" maxlength=""100"" value=""" & ChkString(frs("F_FAQ_QUESTION"),"display") & """></font></td></tr>" & _
            "<tr><td valign=""top""><strong><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>答案:</strong></font></td><td valign=""top"">" & _
            "<textarea name=""Message"" cols=""35"" rows=""7"">" & formatStr(frs("F_FAQ_ANSWER")) & "</textarea>" &_
            "</td></tr><tr><td colspan=2 align=""right""><input type=""submit"" value="" 修改 ""> <input type=""reset"" value="" 重写 "">" & _
			"</form></table>")

        ELSE
        	strFormQ = request.form("fquestion")
            strFormA = request.form("Message")
			
            strFormQ = ChkString(strFormQ,"title")
            strFormA = ChkString(strFormA,"message")
            
            strsql = "UPDATE " & strTablePrefix & "FAQ set F_FAQ_QUESTION='" & strFormQ & "', F_FAQ_ANSWER='" & strFormA & "' WHERE F_ID=" & strFID
			my_conn.execute(strsql)
            response.write("<p><center>问答已经修改完毕, <a href=""admin_faq.asp"">点击这里</a> 返回常见问题管理</center><p>")
        END IF
        %>
        </font>
        <%
    Case "submit"
    	IF request.querystring("do") = "" then
        	%><center>
            <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
            <% if strAllowForumCode = "1" then %>
            * <a href="JavaScript:openWindow3('pop_forum_code.asp')">论坛专用代码</a> 开启<br>
			<% else %>
            * 论坛专用代码 关闭<br>
			<% end if %>
            <% if strAllowHTML = "1" then %>
            * 允许使用 HTML 代码<br>
			<% else %>
            * 禁止使用 HTML 代码<br>
			<% end if %>
            <form method="post" action="admin_faq.asp?action=submit&do=yes" name="PostTopic">
            <select name="font" onChange="thelp(this.options[this.selectedIndex].value)">
				<option value="1">帮助&nbsp;</option>
				<option value="2">完全&nbsp;</option>
				<option selected value="0">基本&nbsp;</option>
	  		  </select>  <a href="Javascript:bold();"><img src="images/icon_editor_bold.gif" width="22" height="22" alt="Bold" border="0"></a><a href="Javascript:italicize();"><img src="images/icon_editor_italicize.gif" width="23" height="22" alt="Italicized" border="0"></a><a href="Javascript:underline();"><img src="images/icon_editor_underline.gif" width="23" height="22" alt="Underline" border="0"></a>
<a href="Javascript:center();"><img src="images/icon_editor_center.gif" width="22" height="22" alt="Centered" border="0"></a>
<a href="Javascript:hyperlink();"><img src="images/icon_editor_url.gif" width="22" height="22" alt="Insert Hyperlink" border="0"></a><a href="Javascript:email();"><img src="images/icon_editor_email.gif" width="23" height="22" alt="Insert Email" border="0"></a><a href="Javascript:image();"><img src="images/icon_editor_image.gif" width="23" height="22" alt="Insert Image" border="0"></a>
<a href="Javascript:showcode();"><img src="images/icon_editor_code.gif" width="22" height="22" alt="Insert Code" border="0"></a><a href="Javascript:quote();"><img src="images/icon_editor_quote.gif" width="23" height="22" alt="Insert Quote" border="0"></a><a href="Javascript:list();"><img src="images/icon_editor_list.gif" width="23" height="22" alt="Insert List" border="0"></a>
<%		if lcase(strIcons) = "1" then %>
<a href="JavaScript:openWindow2('pop_icon_legend.asp')"><img src="images/icon_editor_smilie.gif" width="22" height="22" alt="Insert Smilie" border="0"></a>
<% end if %><br></center>
            <%
        	response.write("<table align=""center"" border=""0"">" & vbcrlf & _
            "<tr><td><strong><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>问题:</strong></td><td><input type=""text"" name=""fquestion"" maxlength=""100""></font></td></tr>" & _
            "<tr><td valign=""top""><strong><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>答案:</strong></font></td><td valign=""top"">" & vbcrlf &  _
            "<textarea name=""Message"" cols=""35"" rows=""7""></textarea>" &vbcrlf & _
            "</td></tr><tr><td colspan=2 align=""right""><input type=""submit"" value="" 提交 ""> <input type=""reset"" value="" 重写 "">" & vbcrlf &  _
			"</table></form>")        
        
        ELSE
        	strFormQ = request.form("fquestion")
            strFormA = request.form("Message")
			
            strFormQ = ChkString(strFormQ,"title")
            strFormA = ChkString(strFormA,"message")
            
            strsql = "INSERT INTO " & strTablePrefix & "FAQ (F_FAQ_QUESTION, F_FAQ_ANSWER)"
            strsql = strsql & " VALUES ('" & strFormQ & "', '" & strFormA & "')"
            my_conn.execute(strsql)        
            response.write("<p><center>添加完成, <a href=""admin_faq.asp"">点击这里</a> 返回常见问题管理</center><p>")        
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
