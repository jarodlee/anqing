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

'#################################################################################
'## Variable declaration 
'#################################################################################
dim strSelectSize
dim intCols, intRows
%>
<!--#INCLUDE FILE="config.asp" --> 
<!--#INCLUDE FILE="inc_functions.asp" -->
<!--#INCLUDE FILE="inc_code.asp" -->
<%
'#################################################################################
'## Initialise variables 
'#################################################################################
strSelectSize = Request.Form("SelectSize") 
strRqMethod = Request.QueryString("method")
strRqForumID = Request.QueryString("FORUM_ID")
strRqCatID = Request.QueryString("CAT_ID")
strCkPassWord = Request.Cookies(strUniqueID & "User")("Pword")
'#################################################################################
'## Page-code start
'#################################################################################
if strSelectSize = "" or IsNull(strSelectSize) then 
	strSelectSize = Request.Cookies(strCookieURL & "strSelectSize")
end if
if not(IsNull(strSelectSize)) then 
	Response.Cookies(strCookieURL & "strSelectSize") = strSelectSize
	Response.Cookies(strCookieURL & "strSelectSize").expires = Now() + 365
end if
%>
<!--#INCLUDE FILE="inc_top.asp" -->
<%

	if strRqForumId <> "" then
	set rs = Server.CreateObject("ADODB.RecordSet")
	strSQL = "SELECT F_PRIVATEFORUMS FROM " & strTablePrefix & "FORUM WHERE FORUM_ID = " & strRqForumID
	set rs = my_conn.execute(strSQL)	
	IsPrivateType =  rs("F_PRIVATEFORUMS")
	set rs = nothing
	end if

select case strSelectSize
	case "1"
		intCols = 45
		intRows = 6
	case "2"
		intCols = 80
		intRows = 12
	case "3"
		intCols = 90
		intRows = 12
	case "4"
		intCols = 130
		intRows = 15
	case else
		intCols = 45
		intRows = 6
end select
%>
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
		try{
			if (document.PostTopic.AuthUsersCombo.options[count].text == "")
			{
				finished = true;
				continue;
			}
		}
		catch(e)
		{
				finished = true;
				return;
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
			try {
			if (document.PostTopic.AuthUsersCombo.options[count].text == "")
			{
				--count;
				continue;
			}
			}
			catch(e)
			{
				return;
			}
			if (document.PostTopic.AuthUsersCombo.options[count].selected )
			{
				for ( z = count ; z < document.PostTopic.AuthUsersCombo.length-1;z++)
				{	
				if ((document.PostTopic.AuthUsers.length-count) == 2)
				{
						document.PostTopic.AuthUsers.options[z].value = "";						
						document.PostTopic.AuthUsers.options[z].text = "";						
				}											
				else					
				{
					document.PostTopic.AuthUsersCombo.options[z].value = document.PostTopic.AuthUsersCombo.options[z+1].value;	
					document.PostTopic.AuthUsersCombo.options[z].text = document.PostTopic.AuthUsersCombo.options[z+1].text;
				}
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
			try{
			if (document.PostTopic.AuthUsers.options[count].text == "")
			{
				finished = true;
				continue;
			}
			}
			catch(e)
			{
				finished = true;
				return;
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
			try{
			if (document.PostTopic.AuthUsers.options[count].text == "")
			{
				--count;
				continue;
			}
			}
			catch(e)
			{
				return;
			}
			if (document.PostTopic.AuthUsers.options[count].selected )
			{
				for ( z = count ; z < document.PostTopic.AuthUsers.length-1;z++)
				{	
					if ((document.PostTopic.AuthUsers.length-count) == 2)
						{
						document.PostTopic.AuthUsers.options[z].value = "";
						document.PostTopic.AuthUsers.options[z].text = "";
						}
						
					else							
						{							
						document.PostTopic.AuthUsers.options[z].value = document.PostTopic.AuthUsers.options[z+1].value;	
						document.PostTopic.AuthUsers.options[z].text = document.PostTopic.AuthUsers.options[z+1].text;
						}
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
		document.PostTopic.SelectSize.value = objform.SelectSize.value;		
		document.cookie =  tmpCookieURL + "strSelectSize=" + objform.SelectSize.value;

}
function OpenPreview()
{
	var curCookie = "strMessagePreview=" + escape(document.PostTopic.Message.value);
	document.cookie = curCookie;
	popupWin = window.open('pop_preview.asp', 'preview_page', 'scrollbars=yes,width=750,height=450')	
}
function OpenPreview2()
{
	var curCookie = "strMessagePreview=" + escape(document.PostTopic.Message.value);
	document.cookie = curCookie;
	popupWin = window.open('pop_spellcheck.asp', 'preview_page', 'scrollbars=yes,width=750,height=450')	
}

//-->
</script>

<% 
if strRqMethod = "EditForum" then
	if (mLev = 4) or (chkForumModerator(strRqForumId, strDBNTUserName) = "1") then
		'## Do Nothing
	else
		Response.Write "<p>错误：只有版主和管理员可以编辑论坛</p>" & vbcrlf
%>
<!--#INCLUDE FILE="inc_footer.asp"-->
<%
		Response.End
	end if
end if

Msg = ""

select case strRqMethod 
	case "Reply"
			if (strNoCookies = 1) or (strDBNTUserName = "")  then
				Msg = Msg & "<b>注意：</b>回复文章之前必须先注册成会员<br>"
				Msg = Msg & "注册请<a href=""policy.asp"">按这里注册</a>。注册是完全免费的！<br>"
			end if
			if (mLev <> 4) and (chkForumModerator(strRqForumId, strDBNTUserName) <> "1") then
			if IsPrivateType = 10 or IsPrivateType = 12 then
				Response.Write "<center><p><b>你不能在此论坛回复主题</b></p></center>"
				Response.End
			end if
			end if
	case "ReplyQuote"
			if (strNoCookies = 1) or (strDBNTUserName = "") then
				Msg = Msg & "<b>注意：</b>回复文章之前必须先注册成会员<br>"
				Msg = Msg & "注册请<a href=""policy.asp"">按这里注册</a>。注册是完全免费的！<br>"
			end if
			if (mLev <> 4) and (chkForumModerator(strRqForumId, strDBNTUserName) <> "1") then
			if IsPrivateType = 10 or IsPrivateType = 12 then
				Response.Write "<center><p><b>你不能在此论坛回复主题</b></p></center>"
				Response.End
			end if
			end if
	case "Topic"
			if (strNoCookies = 1) or (strDBNTUserName = "") then 
				Msg = Msg & "<b>注意：</b> 发表文章之前必须先注册成会员<br>"
				Msg = Msg & "注册请<a href=""policy.asp"">按这里注册</a>。注册是完全免费的！<br>"
			end if
			if (mLev <> 4) and (chkForumModerator(strRqForumId, strDBNTUserName) <> "1") then
			if IsPrivateType = 11 or IsPrivateType = 12 then
				Response.Write "<center><b><p>你不能在本论坛发表新主题</p></b></center>"
				Response.End
			end if
			end if
	case "TopicQuote"
			if (strNoCookies = 1) or (strDBNTUserName = "") then
				Msg = Msg & "<b>注意：</b>回复文章之前必须先注册成会员<br>"
				Msg = Msg & "注册请<a href=""policy.asp"">按这里注册</a>。注册是完全免费的！<br>"
			end if
			if (mLev <> 4) and (chkForumModerator(strRqForumId, strDBNTUserName) <> "1") then
			if IsPrivateType = 11 or IsPrivateType = 12 then
				Response.Write "<center><p><b>你不能在本论坛发表新主题</b></p></center>"
				Response.End
			end if
			end if
	case "Forum"
		Msg = Msg & "<b>注意：</b>只有管理员才能建立新论坛<br>"
	case "URL"
		Msg = Msg & "<b>注意：</b>只有管理员才能建立新连接<br>"
	case "Edit"
		Msg = Msg & "<b>注意：</b>只有作者、版主和管理员可以编辑此文章"
	case "EditTopic"
		Msg = Msg & "<b>注意：</b>只有作者、版主和管理员可以编辑此文章"
	case "EditForum"
		Msg = Msg & "<b>注意：</b>只有版主才能编辑此讯息"
	case "EditCategory"
		Msg = Msg & "注意：只有管理员可以编辑此标题"
end select

if strRqMethod = "Edit" or _
strRqMethod = "ReplyQuote" then
	'## Forum_SQL
	strSql = "SELECT * "
	strSql = strSql & " FROM " & strTablePrefix & "REPLY "
	strSql = strSql & " WHERE " & strTablePrefix & "REPLY.REPLY_ID = " & Request.QueryString("REPLY_ID")

	set rs = my_Conn.Execute (strSql)
	
	strAuthor = rs("R_AUTHOR")

	if strRqMethod = "Edit" then
		TxtMsg = rs("R_MESSAGE")
		if Instr(TxtMsg, "Edited by - ") then TxtMsg = Left(TxtMsg, InstrRev(TxtMsg, "Edited by - ")-36) end if
	else
		if strRqMethod = "ReplyQuote" then
			TxtMsgTemp = rs("R_MESSAGE")
			if Instr(TxtMsgTemp, "Edited by - ") then TxtMsgTemp = Left(TxtMsgTemp, InstrRev(TxtMsgTemp, "Edited by - ")-36) end if
			TxtMsg = "[quote]" & "[b]" & "此文由" & getMemberName(strAuthor) & "发表[/b]" & "。" & vbCrLf & vbCrLf
			TxtMsg = TxtMsg & TxtMsgTemp & "[/quote]" & vbCrLf
		end if
	end if
	if strDBNTUserName = getMemberName(strAuthor) then 
		boolReply =rs("R_MAIL") 
	end if
end if

if strRqMethod = "EditTopic" or _
strRqMethod = "TopicQuote" then
	'## Forum_SQL
	strSql = "SELECT " & strTablePrefix & "TOPICS.CAT_ID, " & strTablePrefix & "TOPICS.FORUM_ID, " & strTablePrefix & "TOPICS.TOPIC_ID, " & strTablePrefix & "TOPICS.T_SUBJECT, " & strTablePrefix & "TOPICS.T_AUTHOR, " & strTablePrefix & "TOPICS.T_MAIL, " & strTablePrefix & "TOPICS.T_MESSAGE " 
	strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
	strSql = strSql & " WHERE " & strTablePrefix & "TOPICS.TOPIC_ID = " & Request.QueryString("TOPIC_ID")

	set rs = my_Conn.Execute (strSql)

	TxtSub = rs("T_SUBJECT")
	strAuthor = rs("T_AUTHOR")

	if strRqMethod = "EditTopic" then
		TxtMsg = rs("T_MESSAGE")
		if Instr(TxtMsg, "Edited by - ") then TxtMsg = Left(TxtMsg, InstrRev(TxtMsg, "Edited by - ")-36) end if
	else
		if strRqMethod = "TopicQuote" then
			TxtMsgTemp = rs("T_MESSAGE")
			if Instr(TxtMsgTemp, "Edited by - ") then TxtMsgTemp = Left(TxtMsgTemp, InstrRev(TxtMsgTemp, "Edited by - ")-36) end if
			TxtMsg = "[quote]" & "[b]" & "此文由" & getMemberName(strAuthor) & "发表[/b]" & "。" & vbCrLf & vbCrLf
			TxtMsg = TxtMsg & TxtMsgTemp & "[/quote]" & vbCrLf
		end if
	end if
	if strDBNTUserName = getMemberName(strAuthor) then boolTopic = rs("T_MAIL")
end if

if strRqMethod = "EditForum" or _
strRqMethod = "EditURL" then
	'## Forum_SQL
'##################
	strSql = "SELECT " & strTablePrefix & "FORUM.F_SUBJECT, " & strTablePrefix & "FORUM.F_URL, " & strTablePrefix & "FORUM.F_DESCRIPTION, " & strTablePrefix & "FORUM.F_PRIVATEFORUMS, " & strTablePrefix & "FORUM.F_PASSWORD_NEW, " & strTablePrefix & "FORUM.F_HIDDEN, " & strTablePrefix & "FORUM.F_URLImage "
'##################
	strSql = strSql & " FROM " & strTablePrefix & "FORUM "
	strSql = strSql & " WHERE " & strTablePrefix & "FORUM.FORUM_ID = " & strRqForumId

	set rs = my_Conn.Execute (strSql)
	
	if strRqMethod = "EditURL" then
		TxtUrl = rs("F_URL")
		TxtUrlImage = rs("F_URLIMAGE")
	end if

	if strRqMethod = "EditForum" or _ 
	strRqMethod = "EditURL" then
		TxtSub = rs("F_SUBJECT")
		TxtMsg = rs("F_DESCRIPTION")
	end if
end if

if strRqMethod = "EditCategory" then

	'## Forum_SQL
	strSql = "SELECT " & strTablePrefix & "CATEGORY.CAT_NAME, "
        strSql = strSql & strTablePrefix & "CATEGORY.CAT_ORDER "
	strSql = strSql & " FROM " & strTablePrefix & "CATEGORY "
	strSql = strSql & " WHERE " & strTablePrefix & "CATEGORY.CAT_ID = " & strRqCatID

	set rs = my_Conn.Execute (strSql)

	if strRqMethod = "EditCategory" then
		TxtSub = rs("CAT_NAME")
		strOrder = rs("CAT_ORDER")
	end if
end if

select case strRqMethod 
	case "Category"
		btn = "建立新分类"
	case "Edit"
		btn = "提交修改"
	case "EditCategory"
		btn = "提交修改"
	case "EditForum"
		btn = "提交修改"
	case "EditTopic"
		btn = "提交修改"
	case "EditURL"
		btn = "提交修改"
	case "Forum"
		btn = "建立新论坛"
	case "Reply"
		btn = "发表回复"
	case "ReplyQuote"
		btn = "发表回复"
	case "TopicQuote"
		btn = "发表回复"
	case "Topic"
		btn = "发表新主题"
	case "URL"
		btn = "建立新连接"
	case else
		btn = "发表"
end select
%>
<table border="0" width="100%" align=center>
  <tr>
    <td width="33%" align="left"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    <img src="<%=strImageURL %>icon_folder_open.gif" border="0">&nbsp;<a href="default.asp">返回论坛首页</a>
<%
if strRqMethod = "EditCategory" then
%>
&nbsp;<img src="<%=strImageURL %>icon_folder_open.gif" border="0">&nbsp;<% =ChkString(Request.QueryString("Cat_Title"),"display") %>
<% 
else 
	if strRqMethod = "Edit" or _
	strRqMethod = "EditTopic" or _
	strRqMethod = "Reply" or _
	strRqMethod = "ReplyQuote" or _
	strRqMethod = "Topic" or _
	strRqMethod = "TopicQuote" then 
%>
&nbsp;<img src="<%=strImageURL %>icon_folder_open.gif" border="0">&nbsp;<a href="FORUM.asp?CAT_ID=<% =strRqCatID %>&FORUM_ID=<% =strRqForumId %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>"><% =ChkString(Request.QueryString("FORUM_Title"),"display") %></a>
<%
	end if 
 end if 

if strRqMethod = "Edit" or _
strRqMethod = "EditTopic" or _
strRqMethod = "Reply" or _
strRqMethod = "ReplyQuote" or _
strRqMethod = "TopicQuote" then 
%>
&nbsp;<img src="<%=strImageURL %>icon_folder_open_topic.gif" border="0">&nbsp;<a href="topic.asp?TOPIC_ID=<% =Request.QueryString("TOPIC_ID") %>&CAT_ID=<% =strRqCatID %>&FORUM_ID=<% =strRqForumId %>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath") %>&Topic_Title=<% =ChkString(left(Request.QueryString("Topic_title"), 50),"urlpath") %>"><% =ChkString(Request.QueryString("Topic_Title"),"title") %></a>
<%
end if 
%>
    </font></td>
  </tr>
</table>

<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
<% =Msg %>
</font></p>
<table border="0" cellspacing="0" cellpadding="0" align=center>
  <tr>
    <td bgcolor="<% =strPopUpBorderColor %>">
    <table border="0" cellspacing="1" cellpadding="4">
<%
if strRqMethod = "Edit" or _
strRqMethod = "EditTopic" or _
strRqMethod = "EditForum" or _
strRqMethod = "Reply" or _
strRqMethod = "ReplyQuote" or _
strRqMethod = "Topic" or _
strRqMethod = "TopicQuote" then 
%>  
  	<form name="Screensize" method="POST" action="post.asp?method=<% =strRqMethod %>&FORUM_ID=<% =strRqForumId%>&CAT_ID=<% =strRqCatID%>&Forum_Title=<% =ChkString(Request.QueryString("FORUM_Title"),"urlpath")%>&TOPIC_ID=<% =ChkString(Request.QueryString("TOPIC_ID"),"urlpath")%>&REPLY_ID=<% =ChkString(Request.QueryString("REPLY_ID"),"urlpath")%>&Topic_Title=<% =ChkString(Request.QueryString("Topic_Title"),"urlpath")%>">
		<input name="Method_Type" type="hidden" value="<% =strRqMethod %>">
		<input name="Type" type="hidden" value="<% =Request.QueryString("type") %>">
		<input name="REPLY_ID" type="hidden" value="<% =Request.QueryString("REPLY_ID") %>">
		<input name="TOPIC_ID" type="hidden" value="<% =Request.QueryString("TOPIC_ID") %>">
		<input name="FORUM_ID" type="hidden" value="<% =strRqForumId %>"> 
		<input name="CAT_ID" type="hidden" value="<% =strRqCatID %>">
		<input name="Author" type="hidden" value="<% =strAuthor %>">
		<input name="Mod_ID" type="hidden" value="<% =Request.QueryString("mod") %>">
		<input name="Cat_Title" type="hidden" value="<% =ChkString(Request.QueryString("Cat_Title"), "hidden") %>">
		<input name="FORUM_Title" type="hidden" value="<% =ChkString(Request.QueryString("FORUM_Title"), "hidden") %>">
		<input name="Topic_Title" type="hidden" value="<% =ChkString(Request.QueryString("TOPIC_Title"), "hidden") %>">
		<input name="M" type="hidden" value="<% =Request.QueryString("M") %>">
		<input name="Refer" type="hidden" value="<% =Request.ServerVariables("HTTP_REFERER") %>">
		<input name="cookies" type="hidden" value="yes">
	<tr>
        <td bgColor="<% =strPopUpTableColor %>" noWrap vAlign="top" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">桌面大小：</font></td>
        <td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
			<select name="SelectSize" size="1" onchange="autoReload(document.Screensize)">  
			<option value="1" <% if strSelectSize = "1" or strSelectSize = "" then Response.Write("SELECTED") %>>640  x 480</option>
			<option value="2" <% if strSelectSize = "2" then Response.Write("SELECTED") %>>800  x 600</option>
			<option value="3" <% if strSelectSize = "3" then Response.Write("SELECTED") %>>1024 x 768</option>
			<option value="4" <% if strSelectSize = "4" then Response.Write("SELECTED") %>>1280 x 1024</option>      
			</select></font>
		</td>
	</tr>
	</form>
<%
end if 
%>
<form action="post_info.asp" method="post" name="PostTopic">
<input name="Method_Type" type="hidden" value="<% =strRqMethod %>">
<input name="Type" type="hidden" value="<% =Request.QueryString("type") %>">
<input name="REPLY_ID" type="hidden" value="<% =Request.QueryString("REPLY_ID") %>">
<input name="TOPIC_ID" type="hidden" value="<% =Request.QueryString("TOPIC_ID") %>">
<input name="FORUM_ID" type="hidden" value="<% =strRqForumId %>"> 
<input name="CAT_ID" type="hidden" value="<% =strRqCatID %>">
<input name="Author" type="hidden" value="<% =strAuthor %>">
<input name="Mod_ID" type="hidden" value="<% =Request.QueryString("mod") %>">
<input name="Cat_Title" type="hidden" value="<% =ChkString(Request.QueryString("Cat_Title"), "hidden") %>">
<input name="FORUM_Title" type="hidden" value="<% =ChkString(Request.QueryString("FORUM_Title"), "hidden") %>">
<input name="Topic_Title" type="hidden" value="<% =ChkString(Request.QueryString("TOPIC_Title"), "hidden") %>">
<input name="M" type="hidden" value="<% =Request.QueryString("M") %>">
<input name="Refer" type="hidden" value="<% =Request.ServerVariables("HTTP_REFERER") %>">
<input name="SelectSize" type="hidden" value="<%= strSelectSize %>">
<input name="cookies" type="hidden" value="yes">


<% 
if mlev = 4 or _
mlev = 3 or _
mlev = 2 or _
mlev = 1 then 
%>
    <input name="UserName" type="hidden" Value="<% =strDBNTUserName%>">
    <input name="Password" type="hidden" value="<% =strCkPassWord%>">
<%
else
	if (lcase(strNoCookies) = "1") or _
	(strDBNTUserName = "" or _
	strCkPassWord = "") then 
%>
      <tr>
        <td bgColor="<% =strPopUpTableColor %>" noWrap vAlign="top" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">用户名：</font></td>
        <td bgColor="<% =strPopUpTableColor %>"><input name="UserName" maxLength="25" size="25" type="text" value="<%=Request.Form("UserName")%>"></td>
      </tr>
      <tr>
        <td bgColor="<% =strPopUpTableColor %>" noWrap vAlign="top" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">密码：</font></td>
        <td bgColor="<% =strPopUpTableColor %>" valign="top"><input name="Password" maxLength="13" size="13" type="password" value="<%=Request.Form("password")%>"></td>
      </tr>
<%
	end if 
end if
 
if strRqMethod = "Edit" or _
strRqMethod = "EditTopic" or _
strRqMethod = "Reply" or _
strRqMethod = "ReplyQuote" or _
strRqMethod = "Topic" or _
strRqMethod = "TopicQuote" then 
%>
<!--#INCLUDE FILE="inc_post_buttons.asp" -->
<% 
end if
if strRqMethod = "Forum" or _
strRqMethod = "URL" or _
strRqMethod = "EditURL" or _
strRqMethod = "EditForum" then 
%>
      <tr>
        <td bgColor="<% =strPopUpTableColor %>" noWrap vAlign="top" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">分类：</font></td>
        <td bgColor="<% =strPopUpTableColor %>">
        <select name="Category" size="1">
<%
'## Forum_SQL
strSql = "SELECT " & strTablePrefix & "CATEGORY.CAT_ID, " & strTablePrefix & "CATEGORY.CAT_NAME "
strSql = strSql & " FROM " & strTablePrefix & "CATEGORY "
if mlev = 3 then 
	strSql = strSql & " WHERE CAT_ID = " & strRqCatID
end if 
strSql = strSql & " ORDER BY " & strTablePrefix & "CATEGORY.CAT_NAME ASC;"

set rsCat = my_Conn.execute (strSql)

'On Error Resume Next
do until rsCat.eof
	Response.Write "          <option value=""" & rsCat("CAT_ID") & """"
	if cint(strRqCatID) = rsCat("CAT_ID") then
		Response.Write " selected"
	end if
	Response.Write ">" & ChkString(rsCat("CAT_NAME"),"display") & "</option>" & vbCrLf
	rsCat.movenext
loop
set rsCat = nothing
%>
        </select>
        </td>
      </tr>
<%
end if

if (strRqMethod = "EditTopic") then
%>
      <tr>
        <td bgColor="<% =strPopUpTableColor %>" noWrap vAlign="top" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">论坛：</font></td>
        <td bgColor="<% =strPopUpTableColor %>">
<%
	if mlev = 3 or _
	mlev = 4 then 
%>
        <select name="Forum" size="1">
<%
	end if 
	'## Forum_SQL
	strSql = "SELECT " & strTablePrefix & "FORUM.CAT_ID, " & strTablePrefix & "FORUM.FORUM_ID, " & strTablePrefix & "FORUM.F_SUBJECT, " & strTablePrefix & "FORUM.F_TYPE "
	strSql = strSql & " FROM " & strTablePrefix & "FORUM "
	strSql = strSql & " WHERE F_TYPE = 0 "
	if mLev = 4 then
		'## Do Nothing
	else
		if mLev = 3 and _
			strMoveTopicMode = "1" then
		else
			if (mLev < 3) then
				strSql = strSql & " AND   FORUM_ID = " & rs("FORUM_ID")
			end if
		end if
	end if
	strSql = strSql & " ORDER BY " & strTablePrefix & "FORUM.F_SUBJECT ASC;"

	set rsForum = my_Conn.execute (strSql)

	if mlev = 3 or _
	mlev = 4 then
		on error resume next
		do until rsForum.eof
			Response.Write "          <option value='" & rsForum("CAT_ID") & "|" & rsForum("FORUM_ID") & "'"
			if cint(strRqForumId) = rsForum("FORUM_ID") then
				Response.Write " selected"
			end if
			Response.Write ">" & ChkString(rsForum("F_SUBJECT"),"display") & "</option>" & vbCrLf
			rsForum.movenext
		loop
	else
		Response.Write "          <font face='" & strDefaultFontFace & "' size=" & strDefaultFontSize & ">" & ChkString(rsForum("F_SUBJECT"),"display") & "</font>"
		Response.Write "<input type='hidden' name='Forum' value='" & rsForum("CAT_ID") & "|" & rsForum("FORUM_ID") & "'>"
	end if
		
	set rsForum = nothing

	if mlev = 3 or _
	mlev = 4 then 
        Response.Write "</select>" & vbcrlf
	end if
end if 

if strRqMethod = "Category" or _
strRqMethod = "EditCategory" or _
strRqMethod = "URL" or _
strRqMethod = "EditURL" or _
strRqMethod = "Forum" or _
strRqMethod = "EditForum" or _
strRqMethod = "EditTopic" or _
strRqMethod = "Topic" then
%>
      <tr>
        <td bgColor="<% =strPopUpTableColor %>" noWrap vAlign="top" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">主题：</font></td>
        <td bgColor="<% =strPopUpTableColor %>"><input maxLength="50" name="Subject" value="<% =Trim(ChkString(TxtSub,"display")) %>" size="40"></td>
      </tr>
<% end if %>

<% if (strRqMethod = "SortCategory") then %>
      <tr>
        <td bgColor="<% =strPopUpTableColor %>" noWrap vAlign="center" align="center" colspan=2><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">设定分类排序
      </tr>
      <tr>
        <td align=center bgcolor="<% =strPopUpTableColor %>" nowrap valign="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">排序</font></td>
        <td align=center bgcolor="<% =strPopUpTableColor %>" nowrap valign="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">分类</font></td>
      </tr>
<%		'## Forum_SQL
		strSql = "SELECT " & strTablePrefix & "CATEGORY.CAT_ID, " 
		strSql = strSql & strTablePrefix & "CATEGORY.CAT_NAME, "
		strSql = strSql & strTablePrefix & "CATEGORY.CAT_ORDER"
		strSql = strSql & " FROM " & strTablePrefix & "CATEGORY "
		strSql = strSql & " ORDER BY " & strTablePrefix & "CATEGORY.CAT_ORDER"
		strSql = strSql & ", " & strTablePrefix & "CATEGORY.CAT_NAME"
		set rslist = Server.CreateObject("ADODB.Recordset")
		rslist.cachesize = 20
		rslist.open  strSql, my_Conn, 3

		If not (rslist.EOF or rslist.BOF) then  '## Forums found in DB for this category
			rslist.movefirst
			rslist.pagesize = 1
			categorycount = cint(rslist.pagecount)
			ordercount = 1 %>
		<input name="NumberCategories" type="hidden" value="<% =categoryCount %>"> 
<%			i = 1
		do while not rslist.eof %>
			<tr><td bgColor="<% =strPopUpTableColor %>">
<%			SelectName = "SortCategory" & ordercount
			SelectID   = "SortCatID" & ordercount %>
			<input name="<% =SelectID %>" type="hidden" value="<% =rslist("CAT_ID") %>"> 
			<SELECT readonly name="<% =SelectName %>">
<%			i = 1
			do while i <= categorycount  %>
				<option value="<% =i %>" <% if (i = rslist("CAT_ORDER")) or (i = ordercount) then Response.Write(" selected") %>><% =i %></option>
<%				i = i + 1
			loop 
			%>
        		</select></td>
        	<td bgColor="<% =strPopUpTableColor %>" noWrap vAlign="top" align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% =ChkString(rslist("CAT_NAME"),"display") %></font></td>
		</tr>
<%		rslist.movenext
		ordercount = ordercount + 1
		loop %>
<%		end if
		set rslist = nothing %>
<% end if %>

<% if (strRqMethod = "EditCategory") then %>
      <tr>
        <td bgColor="<% =strPopUpTableColor %>" noWrap vAlign="top" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">论坛排序：<br></font></td>
	<td bgColor="<% =strPopUpTableColor %>">
<%		'## Forum_SQL
		strSql = "SELECT " & strTablePrefix & "FORUM.FORUM_ID, " 
		strSql = strSql & strTablePrefix & "FORUM.F_SUBJECT, "
		strSql = strSql & strTablePrefix & "FORUM.FORUM_ORDER"
		strSql = strSql & " FROM " & strTablePrefix & "FORUM "
		strSql = strSql & " WHERE Cat_ID = " & Request.QueryString("Cat_ID")
		strSql = strSql & " ORDER BY " & strTablePrefix & "FORUM.FORUM_ORDER"
		strSql = strSql & ", " & strTablePrefix & "FORUM.F_SUBJECT"
		set rslist = Server.CreateObject("ADODB.Recordset")
		rslist.cachesize = 20
		rslist.open  strSql, my_Conn, 3

		If not (rslist.EOF or rslist.BOF) then  '## Forums found in DB for this category
			rslist.movefirst
			rslist.pagesize = 1
			forumcount = cint(rslist.pagecount)
			ordercount = 1 %>
		<input name="NumberForums" type="hidden" value="<% =ForumCount %>"> 
		<table>
<%			i = 1
		do while not rslist.eof %>
			<tr><td bgColor="<% =strPopUpTableColor %>">
<%			SelectName = "SortForum" & ordercount
			SelectID   = "SortForumID" & ordercount %>
			<input name="<% =SelectID %>" type="hidden" value="<% =rslist("FORUM_ID") %>"> 
			<SELECT readonly name="<% =SelectName %>">
<%			i = 1
			do while i <= forumcount  %>
				<option value="<% =i %>" <% if (i = rslist("FORUM_ORDER")) or (i = ordercount) then Response.Write(" selected") %>><% =i %></option>
<%				i = i + 1
			loop 
			%>
        		</select></td>
        	<td bgColor="<% =strPopUpTableColor %>" noWrap vAlign="top" align="left"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% =ChkString(rslist("F_SUBJECT"),"display") %></font></td>
		</tr>
<%		rslist.movenext
		ordercount = ordercount + 1
		loop %>
		</table>
<%		end if
		set rslist = nothing %>
	</td>
      </tr>
<% end if %>

<%
if strRqMethod = "URL" or _
strRqMethod = "EditURL" then 
%>
      <tr>
        <td bgColor="<% =strPopUpTableColor %>" noWrap vAlign="top" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">地址：</font></td>
        <td bgColor="<% =strPopUpTableColor %>"><input maxLength="150" name="Address" value="<% if (TxtURL <> "") then Response.Write(TxtURL) else Response.Write("http://") %>" size="40"></td>
      </tr>
      <tr>
        <td bgColor="<% =strPopUpTableColor %>" noWrap vAlign="top" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">图形连接：</font></td>
        <td bgColor="<% =strPopUpTableColor %>"><input maxLength="150" name="URLImage" value="<% if (TxtURLImage <> "") then Response.Write(TxtURLImage) %>" size="40"></td>
      </tr>
	  
<%
end if
 
if strRqMethod = "Edit" or _
strRqMethod = "URL" or _
strRqMethod = "EditURL" or _
strRqMethod = "Forum" or _
strRqMethod = "EditForum" or _
strRqMethod = "EditTopic" or _
strRqMethod = "Reply" or _
strRqMethod = "ReplyQuote" or _
strRqMethod = "Topic" or _
strRqMethod = "TopicQuote" then 
%>
      <tr>
        <td bgColor="<% =strPopUpTableColor %>" noWrap vAlign="top" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">内容：<br>
        <br>
        <table border=0>
          <tr>
            <td align="left" nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
<% if strAllowHTML = "1" then %>
            * HTML语法开启<br>
<% else %>
            * HTML语法关闭<br>
<% end if %>
<% if strAllowForumCode = "1" then %>
            * <a href="JavaScript:openWindow3('pop_forum_code.asp')">本版专用代码</a>开启<br>
<% else %>
            * 本版专用代码关闭<br>
<% end if %>
            </font>  
            </td>
          </tr>
        </table>
        </font>    
        </td>
        
        <td bgColor="<% =strPopUpTableColor %>"><textarea cols="<%=intCols %>" name="Message" rows="<%=intRows %>" wrap="VIRTUAL"><% =Trim(CleanCode(TxtMsg))%></textarea><br>
        

        </td>
      </tr>
<% 
end if
 
if strRqMethod = "Edit" or _
strRqMethod = "URL" or _
strRqMethod = "EditURL" or _
strRqMethod = "EditTopic" or _
strRqMethod = "Forum" or _
strRqMethod = "Reply" or _
strRqMethod = "ReplyQuote" or _
strRqMethod = "Topic" or _
strRqMethod = "TopicQuote" then 
%>
      <tr>
        <td bgColor="<% =strPopUpTableColor %>">&nbsp;</td>
        <td bgColor="<% =strPopUpTableColor %>">
<% 
	if strRqMethod = "Edit" or _
	strRqMethod = "EditTopic" or _
	strRqMethod = "Reply" or _
	strRqMethod = "ReplyQuote" or _
	strRqMethod = "Topic" or _
	strRqMethod = "TopicQuote" then 
%>
        <table border=0>
         <tr>
          <td>
          <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">模式：
		      <select name="font" onChange="thelp(this.options[this.selectedIndex].value)">
				<option value="1">帮助&nbsp;</option>
				<option selected value="2">完全&nbsp;</option>
				<option value="0">基本&nbsp;</option>
	  		  </select>          
<% 
'#################################################################################
'## File Attachment Stuff
'#################################################################################
		if strRqMethod = "Reply" or _
		strRqMethod = "Edit" or _ 
		strRqMethod = "ReplyQuote" or _
		strRqMethod = "TopicQuote" then

			if intAllowUploads = 1 and STRdbntUserName <> "" then 
				response.write "&nbsp;附加文件：&nbsp;"
				Response.Write "<a href =" & """" & "javascript:openWindow('Mods/pop_upload.asp?TOPIC_ID=" & Request.Querystring("TOPIC_ID") & "&REPLY_ID=" & Request.Querystring("REPLY_ID") & "&MEMBER_ID=" & getMemberID(STRdbntUserName) & "')"" >"
				Response.Write "<img src=""" & strImageURL & "icon_paperclip.gif"" width=23 height=22 border=0 ></a>&nbsp;"
			end if
		end if
'#################################################################################
'## File Attachment Stuff
'#################################################################################
%>
			</font></td>
          </tr>
        </table>
<%
	end if

	if strRqMethod = "Edit" or _
	strRqMethod = "EditTopic" or _
	strRqMethod = "Reply" or _
	strRqMethod = "ReplyQuote" or _
	strRqMethod = "Topic" or _
	strRqMethod = "TopicQuote" then 
		if strRqMethod = "Reply" or _
		strRqMethod = "ReplyQuote" or _
		strRqMethod = "Topic" or _
		strRqMethod = "TopicQuote" then 
%>
        <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
        <input name="Sig" type="checkbox" value="yes" checked>包含个人签名<br>
        </font>
<%
		end if 
		if lcase(strEmail) = "1" then 
			if strRqMethod = "Topic" or _ 
			strRqMethod = "EditTopic" then 
%>  
        <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
        <input name="rmail" type="checkbox" value="1" <%=Chked(boolTopic)%>>当有人回复你的主题时以E-mail通知<br>
        </font>
<%
			else
				if strRqMethod = "Reply" or _
				strRqMethod = "Edit" or _ 
				strRqMethod = "ReplyQuote" or _
				strRqMethod = "TopicQuote" then 
%>
        <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
        <input name="rmail" type="checkbox" value="1" <%=Chked(boolReply)%>>当有新回复时以E-mail通知<br>
        </font>
<%
				end if
			end if
		end if

		if 	((mLev = 4) or (chkForumModerator(strRqForumId, strDBNTUserName) = "1")) _
		and (strRqMethod = "Topic" or strRqMethod = "Reply" or _
		strRqMethod = "ReplyQuote" or strRqMethod = "TopicQuote") then %>
        <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
        <input name="lock" type="checkbox" value="1">发表后锁定本主题<br>
        </font>
<%
		end if
	end if
end if 
%>
        </td>
      </tr>
      <tr>
        <td bgColor="<% =strPopUpTableColor %>">&nbsp;</td>
        <td bgColor="<% =strPopUpTableColor %>"><input name="Submit" type="submit" value="<% =btn %>">
<%
		if strRqMethod = "Reply" or _
		strRqMethod = "Edit" or _ 
		strRqMethod = "ReplyQuote" or _
		strRqMethod = "Topic" or _
		strRqMethod = "TopicQuote" then 
%>      
		&nbsp;<input name="Preview" type="button" value=" 预览 " onclick="OpenPreview()">
<% 
' ##########################  SpellCheck MOD  ############################## 
	if intSpellCheck then %>
		&nbsp;<input name="Preview" type="button" value=" 拼字检查（英文）" onclick="OpenPreview2()">
<% 	End if 
' ##########################  SpellCheck MOD  ############################## %>

<%
		end if
%>  
        &nbsp;<input name="Reset" type="reset" value="全部清除"></td>
      </tr>
<% 
if strPrivateForums <> "0" then 
	if strRqMethod = "Forum" or _
	strRqMethod = "URL" or _
	strRqMethod = "EditURL" or _
	strRqMethod = "EditForum" then 
		if strRqMethod = "EditForum" or _
		strRqMethod = "EditURL" then
			ForumAuthType = rs("F_PRIVATEFORUMS")
			ForumHide = rs("F_HIDDEN")
		else
			ForumAuthType = 0
			ForumHide = 0
		end if
%>
      <tr>
        <td bgColor="<% =strPopUpTableColor %>" noWrap vAlign="top" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">授权模式：</font></td>
        <td bgColor="<% =strPopUpTableColor %>"><SELECT readonly name="AuthType">
        <option value="0" <% if ForumAuthType = 0 then Response.Write(" selected") %>>任何人（包括游客跟会员）</option>
        <option value="4" <% if ForumAuthType = 4 then Response.Write(" selected") %>>只有会员</option>
        <option value="5" <% if ForumAuthType = 5 then Response.Write(" selected") %>>只有会员（隐藏）</option>
<%
		if strRqMethod = "Forum" or _
		strRqMethod = "EditForum" then 
%>
        <option value="2" <% if ForumAuthType = 2 then Response.Write(" selected") %>>密码登陆</option>
        <option value="7" <% if ForumAuthType = 7 then Response.Write(" selected") %>>只有会员＆密码保护</option>
        <option value="3" <% if ForumAuthType = 3 then Response.Write(" selected") %>>授权的会员＆密码登陆</option>
<%		end if %>
        <option value="1" <% if ForumAuthType = 1 then Response.Write(" selected") %>>授权的会员</option>
        <option value="6" <% if ForumAuthType = 6 then Response.Write(" selected") %>>授权的会员（隐藏）</option>
<%	If strNTGroups = "1" then %>	
        <option value="9" <% if ForumAuthType = 9 then Response.Write(" selected") %>>NT Global Group</option>
        <option value="8" <% if ForumAuthType = 8 then Response.Write(" selected") %>>NT Global Group（隐藏）</option>
<%	end if%>
<%'############### READ/WRITE ACCESS ###########################%>
        <option value="10" <% if ForumAuthType = 10 then Response.Write(" selected") %>>只有版主和管理员可以回复</option>				
        <option value="11" <% if ForumAuthType = 11 then Response.Write(" selected") %>>任何会员都可以回复</option>
        <option value="12" <% if ForumAuthType = 12 then Response.Write(" selected") %>>只有版主和管理员可以发表及回复</option>		
<%'############### READ/WRITE ACCESS ###########################%>
        </select>
<%'#################%>
		&nbsp;<input type="checkbox" name="HideForum" value="1" <% if ForumHide = 1 then Response.Write(" checked") %>>隐藏
<%'#################%>
<%
		if strRqMethod = "Forum" or _
		strRqMethod = "EditForum" then 
			if strRqMethod = "EditForum" then 
				If rs("F_PASSWORD_NEW") <> " " Then
					strPassword = ChkString(rs("F_PASSWORD_NEW"),"password")
				else 
					strPassword = " "
				end if
			else
				strPassword = " "
			end if
%>
        <br><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">密码<% if strNTGroups = "1" then Response.Write("或 Global Groups") %>：<br></font> 
        <input maxLength="255" type="text" name="AuthPassword" size="60" value="<%=strPassword%>">
<%
		end if 
%>
        </td>
      </tr>
      <tr>
        <td bgColor="<% =strPopUpTableColor %>" noWrap vAlign="top" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">授权会员名单：</font></td>
<%
			'#################################################################################
			'## Allowed User - listbox Code
			'#################################################################################
		if strRqMethod = "EditForum" or strRqMethod = "EditURL" then
			strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME "
			strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
			strSql = strSql & " WHERE NOT EXISTS (SELECT *"
			strSql = strSql & " FROM " & strTablePrefix & "ALLOWED_MEMBERS "
			strSql = strSql & " WHERE " & strTablePrefix & "ALLOWED_MEMBERS.FORUM_ID = " & strRqForumID 
			strSql = strSql & " AND " & strTablePrefix & "ALLOWED_MEMBERS.MEMBER_ID = " & strMemberTablePrefix & "MEMBERS.MEMBER_ID) "
		else
			strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME "
			strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
		end if
		strSql = strSql & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_NAME ASC;"
		on error resume next
		set rsMember = my_Conn.execute (strSql)

		strSql = "SELECT " & strTablePrefix & "ALLOWED_MEMBERS.MEMBER_ID "
		strSql = strSql & " FROM " & strTablePrefix & "ALLOWED_MEMBERS "
		strSql = strSql & " WHERE " & strTablePrefix & "ALLOWED_MEMBERS.FORUM_ID = " & strRqForumID

		set rsAllowedMember = my_Conn.execute (strSql)

		tmpStrUserList  = ""
		if strRqMethod = "EditForum" or strRqMethod = "EditURL" then
		do while not (rsAllowedMember.EOF or rsAllowedMember.BOF)
			if tmpStrUserList = "" then
				tmpStrUserList = rsAllowedMember("MEMBER_ID")
			else
				tmpStrUserList = tmpStrUserList & "," & rsAllowedMember("MEMBER_ID")
			end if
			rsAllowedMember.movenext
		loop
		end if
		uSelectSize = 10
%>                   
<td bgColor="<% =strPopUpTableColor %>">
	<table><tr>
		<td>
			<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">论坛成员：</font><br>
			<select name="AuthUsersCombo" size="<%=uSelectSize %>" multiple >
<%
	'## Pick from list
	rsMember.movefirst
	do until rsMember.eof
		if not(Instr(tmpStrUserList ,rsMember("MEMBER_ID")) > 0) then
			Response.Write "<option value=""" & rsMember("MEMBER_ID") & """" & isSel & ">" & ChkString(rsMember("M_NAME"),"display") & "</option>" & vbCrLf
		end if
		rsMember.movenext
	loop
%>      
	    	<option value="<% if tmpStrUserList <> "" then Response.Write tmpStrUserList end if %>"></option>
		</select>
	</td>
	<td width="15" align="center" valign="middle">
		<a href="javascript:MoveWholeList('Del')"><img src="<%=strImageURL %>icon_Private_remall.gif" width="23" height="22" border="0" alt=""></a>
		<a href="javascript:InsertSelection('Del')"><img src="<%=strImageURL %>icon_Private_remove.gif" width="23" height="22" border="0" alt=""></a>
		<a href="javascript:InsertSelection('Add')"><img src="<%=strImageURL %>icon_Private_add.gif" width="23" height="22" border="0" alt=""></a>
		<a href="javascript:MoveWholeList('Add')"><img src="<%=strImageURL %>icon_Private_addall.gif" width="23" height="22" border="0" alt=""></a>
	</td>
	<td>
		<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">授权会员：</font><br>
		<select name="AuthUsers" size="<%=uSelectSize %>" multiple >
<%
	'## Selected List
	rsAllowedMember.movefirst
	if strRqMethod = "EditForum" or strRqMethod = "EditURL" then	
	do until rsAllowedMember.EOF
		Response.Write "          <option value=""" & rsAllowedMember("MEMBER_ID") & """>" & ChkString(getMemberName(rsAllowedMember("MEMBER_ID")),"display") & "</option>" & vbCrLf
		rsAllowedMember.movenext
	loop
	end if
	set rsAllowedMember = nothing
%>
        	<option value="<% if tmpStrUserList <> "" then Response.Write tmpStrUserList end if %>"></option>
		</select>
    </td>
  </tr>
</table>
	</td>
      </tr>
      <tr>
        <td bgColor="<% =strPopUpTableColor %>">&nbsp;</td>
        <td bgColor="<% =strPopUpTableColor %>"><input name="Submit" type="submit" value="<% =btn %>" onclick="selectUsers()">&nbsp;<input name="Reset" type="reset" value="全部清除"></td>
      </tr>
<%
		'#################################################################################
		'## Allowed User - End of listbox code
		'#################################################################################
	end if
end if 
%>
    </table>
    </td>
  </tr>
</table>
</form>

<%
if strRqMethod = "Reply" or _
strRqMethod = "TopicQuote" or _
strRqMethod = "ReplyQuote" then
%>
<table border="0" width="95%" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td bgcolor="<% =strTableBorderColor %>">
    <table border="0" width="100%" cellspacing="1" cellpadding="4">
      <tr>
        <td bgcolor="<% =strHeadCellColor %>" colspan="2" align="center"><b><font <% =strDefaultFontFace %> size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">原  &nbsp;&nbsp;&nbsp; 文</font></b></td>
      </tr>
<%
	'## Forum_SQL
	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strTablePrefix & "TOPICS.T_MESSAGE " 
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS, " & strTablePrefix & "TOPICS "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "TOPICS.T_AUTHOR AND "
	strSql = strSql & "       " & strTablePrefix & "TOPICS.TOPIC_ID = " &  Request.QueryString("TOPIC_ID")

	set rs = my_Conn.Execute (strSql) 

	Response.Write "      <tr>" & vbCrLf
	Response.Write "        <td bgcolor='" & strForumFirstCellColor & "' valign=top width='" & strTopicWidthLeft & "'"
	if lcase(strTopicNoWrapLeft) = "1" then
		Response.Write " nowrap"
	end if 
	Response.Write "><font color='" & strForumFontColor & "' face='" & strDefaultFontFace & "' size='2'><b>" & ChkString(rs("M_NAME"),"display") & "</b></font></td>" & vbCrLf
	Response.Write "        <td bgcolor='" & strForumCellColor & "' valign='top' width='" & strTopicWidthRight & "'" 
	if lcase(strTopicNoWrapRight) = "1" then
		Response.Write " nowrap"
	end if 
	Response.Write "><font color='" & strForumFontColor & "' face='" & strDefaultFontFace & "' size='2'>" & formatStr(rs("T_MESSAGE")) & "</font></td>" & vbCrLf
	Response.Write "      </tr>" & vbCrLf

	'## Forum_SQL - Get all replies to Topic from the DB
	strSql ="SELECT " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strTablePrefix & "REPLY.R_MESSAGE "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS, " & strTablePrefix & "REPLY "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & strTablePrefix & "REPLY.R_AUTHOR "
	strSql = strSql & " AND   TOPIC_ID = " & Request.QueryString("TOPIC_ID") & " "
	strSql = strSql & " ORDER BY " & strTablePrefix & "REPLY.R_DATE"

	set rs = Server.CreateObject("ADODB.Recordset")

'	rs.cachesize=15
	rs.open  strSql, my_Conn, 3

	strI = 0 
	if rs.EOF or rs.BOF then
		Response.Write ""
	else
		rs.movefirst
		do until rs.EOF
			if strI = 0 then
	 			CColor = strAltForumCellColor
			else
				CColor = strForumCellColor
			end if
			Response.Write "      <TR>" & vbCrLf   & _
						   "        <TD bgcolor='" & CColor & "' valign='top'"
			if lcase(strTopicNoWrapLeft) = "1" then
				Response.Write " nowrap"
			end if 
			Response.Write "><b><font color='" & strForumFontColor & "' face='" & strDefaultFontFace & "' size='2'>" &  ChkString(rs("M_NAME"),"display") & "</font></b></td>" & vbCrLf & _
						   "        <TD bgcolor='" & CColor & "' valign='top'"
			if lcase(strTopicNoWrapRight) = "1" then
				Response.Write " nowrap"
			end if
			Response.Write "><font color='" & strForumFontColor & "' face='" & strDefaultFontFace & "' size='2'>" & formatStr(rs("R_MESSAGE")) & "</font></td>" & vbCrLf & _
						   "      </tr>" & vbCrLf
			rs.MoveNext
			strI = strI + 1
			if strI = 2 then 
				strI = 0
			end if
		loop
	end if

	Response.Write	"    </table>" & vbCrLf  & _
					"    </td>" & vbCrLf & _
					"  </tr>" & vbCrLf  & _
					"</table>" & vbCrLf & _
					"</font>" & vbCrLf  
end if
%> 
<!--#INCLUDE FILE="inc_footer.asp" -->
