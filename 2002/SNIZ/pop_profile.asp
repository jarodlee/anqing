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
<!--#INCLUDE FILE="inc_functions.asp" -->
<%

if not(strUseExtendedProfile) then
%>
<!--#INCLUDE FILE="inc_top_short.asp" -->
<%
else
%>
<!--#INCLUDE FILE="inc_top.asp" -->
<% end if 
if Request.QueryString("mode") = "Moderator" then
select case Request.QueryString("action")
case "del"
	strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " SET M_LEVEL = 1 "
	strSql = strSql & " WHERE MEMBER_ID = " & Request.QueryString("ID")
	my_Conn.Execute(strsql)
case "add"
	strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " SET M_LEVEL = 2 "
	strSql = strSql & " WHERE MEMBER_ID = " & Request.QueryString("ID")
	my_Conn.Execute(strsql)
end select %>
<meta http-equiv="Refresh" content="0; URL=members.asp">
<% end if
if strAuthType = "nt" then
	if ChkAccountReg() <> "1" then %>
<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
<b>注意：</b>此用户帐号尚未注册所以没有会员资料。<br>
如果这是你的帐号，请<a href="register.asp?mode=Register">按这里</a>注册。
</font></p>		
<!--#INCLUDE FILE="inc_footer.asp" -->
<% 
		Response.End 
	end if
end if
%>
<%
select case Request.QueryString("mode") 

	case "display" '## Display Profile

		'## Forum_SQL
		strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME " 
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_USERNAME, " & strMemberTablePrefix & "MEMBERS.M_EMAIL "
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_FIRSTNAME, " & strMemberTablePrefix & "MEMBERS.M_LASTNAME " 
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_TITLE"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_PASSWORD"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_ICQ"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_YAHOO"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_AIM"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_COUNTRY "
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_POSTS"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_CITY " 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_STATE " 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_COUNTRY " 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_POSTS " 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_HIDE_EMAIL " 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_DATE " 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_PHOTO_URL " 
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_HOMEPAGE" 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_LINK1, " & strMemberTablePrefix & "MEMBERS.M_LINK2 "
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_AGE, " & strMemberTablePrefix & "MEMBERS.M_MARSTATUS " 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_SEX, " & strMemberTablePrefix & "MEMBERS.M_OCCUPATION " 
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_SIG"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_HOBBIES, " & strMemberTablePrefix & "MEMBERS.M_LNEWS " 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_QUOTE, " & strMemberTablePrefix & "MEMBERS.M_BIO "
		strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " WHERE MEMBER_ID=" & Request.QueryString("id")

		set rs = my_Conn.Execute(strSql)
		
		strMyHobbies = rs("M_HOBBIES")
		strMyLNews = rs("M_LNEWS")
		strMyQuote = rs("M_QUOTE")
		strMyBio = rs("M_BIO")
		Set HideMail = rs("M_HIDE_EMAIL")
	
		if strUseExtendedProfile then
			strColspan = " colspan=""2"" "
			strIMURL1 = "javascript:openWindow('"
			strIMURL2 = "')"
		else
			strColspan = ""
			strIMURL1 = ""
			strIMURL2 = ""
		end if
		mLev = cint(ChkUser2(Request.Cookies(strUniqueID & "User")("Name"), Request.Cookies(strUniqueID & "User")("Pword")))
%>

	<table border="0" width="100%" cellspacing="0" cellpadding="0" bgColor="<% =strPageBGColor %>">
	  <tr>
	    <td bgColor="<% =strPageBGColor %>" align=center  <%= strColspan %>>
	    <p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">用户资料<br></font></p>
	    </td>
	  </tr>
	  <tr>
	  	 <td bgColor="<% =strPageBGColor %>" align=center <%= strColspan %>>
	  		<table border="0" width="95%" cellspacing="0" cellpadding="4" align=center>
	  		<tr>
<% if mLev = 4 then %>
		<td valign=top align=left bgcolor="<% =strHeadCellColor %>">&nbsp;<a href="pop_profile.asp?mode=Modify&ID=<% =rs("MEMBER_ID") %>&name=<% =ChkString(rs("M_NAME"),"urlpath") %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>"><b><% =ChkString(rs("M_NAME"),"display") %></a></b></font></td>
<% else %>				
				<td valign=top align=left bgcolor="<% =strHeadCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>"><b>&nbsp;<% =ChkString(rs("M_NAME"),"display")  %></b></font></td>
<% end if%>
				<td valign=top align=right bgcolor="<% =strHeadCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">注册日期：<% =ChkDate(rs("M_DATE")) %>&nbsp;</font></td>
			</tr>
			</table>
		</td>
	  </tr>
	  <tr>
	    <td bgcolor="<% =strPageBGColor %>" align=left valign=top>
			<table border="0" width="95%" cellspacing="1" cellpadding="0" align="center" bgcolor="<% =strPageBGColor %>">
				<tr bgcolor="<% =strPageBGColor %>">
<%
			if strUseExtendedProfile then
%>
				  <td width="40%" bgColor="<% =strPageBGColor %>" valign=top>
				    <table border="0" width="100%" cellspacing="1" cellpadding="3" bgcolor="<% =strTableBorderColor %>">								    
<%					if strPicture = "1" then %>
					 <tr>
						<td align=center bgcolor="<% =strCategoryCellColor %>" colspan="2"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">&nbsp;个人头像：</font></b></td>
					 </tr>
					 <tr>
 					    <td bgColor=<% =strPopUpTableColor %> align=center colspan="2">
<%							if Trim(rs("M_PHOTO_URL")) <> "" and lcase(rs("M_PHOTO_URL")) <> "http://" then %>
								<a href="<% =ChkString(rs("M_PHOTO_URL"), "display") %>"><img src="<% =rs("M_PHOTO_URL") %>" width="150" height="150" border=0 hspace="2" vspace="2"></a><br><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">点击浏览全图</font>
<%							else %>
								<img src="no_photo.gif" alt="没有设定图片" width="150" height="150" border=0 hspace="2" vspace="2"></a>
<%							end if %>
						</td>
					</tr>
<%					end if ' strPicture %>					
					 <tr>
						<td align=center bgcolor="<% =strCategoryCellColor %>" colspan="2"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">&nbsp;联系资料&nbsp;</font></b></td>
					 </tr>
					 <tr>
						<td bgColor="<% =strPopUpTableColor %>" align=right width="10%" nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"> 电子邮件：&nbsp;</font></b></td>
<% if(HideMail = 0) then %>
						<td bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">&nbsp;<a href="JavaScript:openWindow('pop_mail.asp?id=<% =rs("MEMBER_ID") %>')"><% =ChkString(rs("M_EMAIL"), "display") %></a>&nbsp;</font></td>
<% else %>
<td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"> 地址隐藏  </font></td>
<% end if %>
					</tr>
<%				if strICQ = "1" then
					if Trim(rs("M_ICQ")) <> "" then %>
					<tr>
						<td bgColor=<% =strPopUpTableColor %> align=right nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">ICQ号码：&nbsp;</font></b></td>
						<td bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="<% =strIMURL1 %>pop_messengers.asp?mode=ICQ&ICQ=<% =ChkString(rs("M_ICQ"), "urlpath") %>&M_NAME=<% =ChkString(rs("M_NAME"), "urlpath") %><% =strIMURL2 %>"><img src="http://online.mirabilis.com/scripts/online.dll?icq=<% =ChkString(rs("M_ICQ"), "urlpath") %>&img=5" border="0" align="absmiddle"><% =ChkString(rs("M_ICQ"), "urlpath") %></a></font></td>
					</tr>
<%					end if
				end if
				if strYAHOO = "1" then
					if Trim(rs("M_YAHOO")) <> "" then %>
					<tr>
						<td bgColor=<% =strPopUpTableColor %> align=right nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">OICQ号码：&nbsp;</font></b></td>
						<td bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="http://search.tencent.com/cgi-bin/friend/user_show_info?ln=<% =rs("M_YAHOO") %>"  target=_blank><img align=absmiddle width=16 height=16 src="http://icon.tencent.com/<% =rs("M_YAHOO") %>/s/00/" border=0><% =rs("M_YAHOO") %></a></font></td>
					</tr>
<%					end if
				end if
				if strAIM = "1" then
					if Trim(rs("M_AIM")) <> "" then %>
					<tr>
						<td bgColor=<% =strPopUpTableColor %> align=right nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">AIM号码：&nbsp;</font></b></td>
						<td bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="<% =strIMURL1 %>pop_messengers.asp?mode=AIM&AIM=<% =ChkString(rs("M_AIM"), "urlpath") %>&M_NAME=<% =ChkString(rs("M_NAME"), "urlpath") %><% =strIMURL2 %>"><% =ChkString(rs("M_AIM"), "urlpath") %></a>&nbsp;</font></td>
					</tr>
<%					end if
				end if 

				if strRecentTopics = "1" then
					strStartDate = DateToStr(dateadd("d", -30, strForumTimeAdjust))
					'## Forum_SQL - Find all records for the member
					strsql = "SELECT " & strTablePrefix & "FORUM.FORUM_ID, " & strTablePrefix & "FORUM.F_SUBJECT, " & strTablePrefix & "FORUM.CAT_ID, " & strTablePrefix & "TOPICS.TOPIC_ID, " 
					strsql = strsql & strTablePrefix & "TOPICS.T_SUBJECT, " & strTablePrefix & "TOPICS.T_STATUS,  " & strTablePrefix & "TOPICS.T_LAST_POST,  " & strTablePrefix & "TOPICS.T_REPLIES  "
					strsql = strsql & " FROM ((" & strTablePrefix & "FORUM LEFT JOIN " & strTablePrefix & "TOPICS "
					strsql = strsql & " ON " & strTablePrefix & "FORUM.FORUM_ID = " & strTablePrefix & "TOPICS.FORUM_ID) LEFT JOIN " & strTablePrefix & "REPLY "
					strsql = strsql & " ON " & strTablePrefix & "TOPICS.TOPIC_ID = " & strTablePrefix & "REPLY.TOPIC_ID) "
					strsql = strsql & " WHERE (T_DATE > '" & strStartDate & "') "
					strsql = strsql & " AND (" & strTablePrefix & "TOPICS.T_AUTHOR = " & Request.QueryString("id") & " "
					strsql = strsql & " OR " & strTablePrefix & "REPLY.R_AUTHOR = " & Request.QueryString("id") & ")"
					strSql = strSql & " ORDER BY " & strTablePrefix & "TOPICS.T_LAST_POST DESC, " & strTablePrefix & "TOPICS.TOPIC_ID DESC"
						
					set rs2 = my_Conn.Execute(strsql)					
%>					    
					 <tr>
						<td align=center bgcolor="<% =strCategoryCellColor %>" colspan="2"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">最近发表的主题</font></b></td>
					 </tr>

<%					if rs2.EOF or rs2.BOF then  %>
					  <tr>
						<td bgcolor="<% =strPopUpTableColor %>" colspan="2"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">&nbsp;<br>&nbsp;<b>没有找到相符的资料<br>&nbsp;</b></font></td>
					  </tr>
<%
					else 
						currTopic = 0
						TopicCount = 0
	%>
					<tr>
						<td bgColor=<% =strPopUpTableColor %> valign=top colspan="2">
							<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<%				
						do until rs2.EOF or (TopicCount = 10)
							if chkDisplayForum(rs2("FORUM_ID")) then 
								if currTopic <> rs2("TOPIC_ID") then %>
									<tr>
										<td bgcolor="<% = strPopUpTableColor %>">
										<a href="topic.asp?TOPIC_ID=<% =rs2("TOPIC_ID") %>&FORUM_ID=<% =rs2("FORUM_ID") %>&CAT_ID=<% =rs2("CAT_ID") %>&Topic_Title=<% =ChkString(left(rs2("T_SUBJECT"), 50),"urlpath") %>&Forum_Title=<% =ChkString(rs2("F_SUBJECT"),"urlpath") %>">
<%									if rs2("T_STATUS") <> 0 then
										if strHotTopic = "1" then
											if rs2("T_LAST_POST") > Session(strCookieURL & "last_here_date") then
												if rs2("T_REPLIES") >= intHotTopicNum then%>
				        				<img src=images/icon_folder_new_hot.gif height=15 width=15 alt=热门主题 border=0></a></td>
<%												else%>
			    						<img src=images/icon_folder_new.gif height=15 width=15 alt=新主题 border=0></a></td>
<%												end if
											else
												if rs2("T_REPLIES") >= intHotTopicNum then%>
				        				<img src=images/icon_folder_hot.gif height=15 width=15 alt=热门主题 border=0></a></td>
<%												else%>
					    				<img src=images/icon_folder.gif height=15 width=15 border=0></a></td>
<%												end if
											end if
										else
											if rs2("T_LAST_POST") > Session(strCookieURL & "last_here_date") then %>
										<img src=images/icon_folder_new.gif height=15 width=15 alt=新主题 border=0></a></td>
<%											else%>
										<img src=images/icon_folder.gif height=15 width=15 border=0></a></td> 
<%											end if
										end if
									else 
										if rs2("T_LAST_POST") > Session(strCookieURL & "last_here_date") then %>
										<img src="images/icon_folder_new_locked.gif" alt="主题锁定" border="0"></a></td>
<%										else %>
										<img src="images/icon_folder_locked.gif" alt="主题锁定" border="0"></a></td>
<%										end if
									end if %>
										<td bgcolor="<% = strPopUpTableColor %>" valign="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="topic.asp?TOPIC_ID=<% =rs2("TOPIC_ID") %>&FORUM_ID=<% =rs2("FORUM_ID") %>&CAT_ID=<% =rs2("CAT_ID") %>&Topic_Title=<% =ChkString(left(rs2("T_SUBJECT"), 50),"urlpath") %>&Forum_Title=<% =ChkString(rs2("F_SUBJECT"),"urlpath") %>"><% =ChkString(left(rs2("T_SUBJECT"), 50),"display") %></a>&nbsp;</font></td>
									</tr>
<%
									TopicCount = TopicCount + 1
								end if 
								currTopic = rs2("TOPIC_ID")
							end if
							rs2.MoveNext 
							TopicCount = TopicCount + 1
						loop 
%>							</table>
						</td>
					</tr>				
<%					end if 
					rs2.close
					set rs2 = nothing

				elseif (strHomepage + strFavLinks) > 0 and (strRecentTopics = "0") then  %>
				<tr>
					<td align=center bgcolor="<% =strCategoryCellColor %>" colspan="2">
					<b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">相关连接&nbsp;</font></b></td>			
<%					if strHomepage = "1" then %>
						<tr>
							<td bgColor=<% =strPopUpTableColor %> align=right nowrap width="10%"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">个人主页：&nbsp;</font></b></td>
<%							if Trim(rs("M_HOMEPAGE")) <> "" and lcase(trim(rs("M_HOMEPAGE"))) <> "http://" and Trim(lcase(rs("M_HOMEPAGE"))) <> "https://" then %>
							<td bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="<% =rs("M_Homepage") %>" target="_blank"><% =rs("M_Homepage") %></a>&nbsp;</font></td>
<%							else %>
							<td bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">没有个人主页</font></td>
<%							end if %>
						</tr>
<%					end if
						
					if strFavLinks = "1" then %>
						<tr>
							<td bgColor=<% =strPopUpTableColor %> align=right nowrap width="10%"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">推荐连接：&nbsp;</font></b></td>
<%						if Trim(rs("M_LINK1")) <> "" and lcase(trim(rs("M_LINK1"))) <> "http://" and Trim(lcase(rs("M_LINK1"))) <> "https://" then %>
							<td bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="<% =rs("M_LINK1") %>" target="_Blank"><% =rs("M_LINK1") %></a>&nbsp;</font></td>
<%						else %>
							<td bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">没有推荐连接</font></td>
						</tr>
<%							if Trim(rs("M_LINK2")) <> "" and lcase(trim(rs("M_LINK2"))) <> "http://" and Trim(lcase(rs("M_LINK2"))) <> "https://" then %>
						<tr>
							<td bgColor=<% =strPopUpTableColor %> align=right nowrap width="10%"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">&nbsp;</font></b></td>
							<td bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="<% =rs("M_LINK2") %>" target="_Blank"><% =rs("M_LINK2") %></a>&nbsp;</font></td>

<%							end if
						end if %>
						</tr>		
<%					end if 

				end if ' strRecentTopics
%>	
					</table>	
				</td>
				<td valign=top width="3%" bgColor=<% =strPageBGColor %> >
				&nbsp;
				</td>
<%
		end if ' UseExtendedMemberProfile
%>
				<td bgColor=<% =strPageBGColor %> valign=top>
					<table border="0" width="100%" cellspacing="1" cellpadding="3" valign=top bgcolor="<% =strTableBorderColor %>">
						<tr>
							<td valign="top" align=center colspan="2" bgcolor="<% =strCategoryCellColor %>"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">基本资料</font></b></td>
						</tr>
						<tr>	
							<td bgColor=<% =strPopUpTableColor %> align=right nowrap width="10%" valign=top><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"> 会员名字：&nbsp;</font></b></td>
							<td bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% =ChkString(rs("M_NAME"),"display") %>&nbsp;</font></td>					  
						</tr>
<%				if strAuthType = "nt" then %>
						<tr>
							<td bgColor="<% =strPopUpTableColor %>" align=right nowrap valign=top><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">你的帐号：&nbsp;</font></b></td>
							<td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSizederFontSize %>"><%= ChkString(rs("M_USERNAME"),"display") %></FONT></td>
						</tr>
<%				end if 
				if strFullName = "1" and (Trim(rs("M_FIRSTNAME")) <> ""  or  Trim(rs("M_LASTNAME")) <> "" ) then
%>
						<tr>
							<td bgColor="<% =strPopUpTableColor %>" align=right nowrap valign=top><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">真实姓名：</font></b></td>
							<td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% =ChkString(rs("M_LASTNAME"), "display") %>&nbsp;<% =ChkString(rs("M_FIRSTNAME"), "display") %></font></td>
						</tr>
<%
				end if
				if (strCity = "1" and Trim(rs("M_CITY")) <> "") or (strCountry = "1" and Trim(rs("M_COUNTRY")) <> "") or (strCountry = "1" and Trim(rs("M_STATE")) <> "") then
%>			
						<tr>
							<td bgColor="<% =strPopUpTableColor %>" align=right nowrap valign=top><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">国家地区：</font></b></td>
							<td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
<%							Response.Write(rs("M_CITY")) 
							if Trim(rs("M_CITY")) <> "" then
								Response.Write("&nbsp;")
							end if
							if Trim(rs("M_STATE")) <> "" then
								Response.Write(ChkString(rs("M_STATE"), "display") & "<br>")
							end if
							Response.Write(ChkString(rs("M_COUNTRY"), "display")) 
%>
							</font>
							</td>
						</tr>
<%
				end if
				if (strAge = "1" and Trim(rs("M_AGE")) <> "") then
%>							
						<tr>
							<td bgColor="<% =strPopUpTableColor %>" align=right nowrap valign=top><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">年&nbsp;&nbsp;&nbsp;&nbsp;龄：&nbsp;</font></b></td>
							<td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% =ChkString(rs("M_AGE"), "display") %></font></td>
						</tr>
<%
				end if
				if (strMarStatus = "1" and Trim(rs("M_MARSTATUS")) <> "") then
%>			
						<tr>
							<td bgColor="<% =strPopUpTableColor %>" align=right nowrap valign=top><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">婚姻状况：&nbsp;</font></b></td>
							<td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% = ChkString(rs("M_MARSTATUS"), "display") %></font></td>
						</tr>
<%
				end if
				if (strSex = "1" and Trim(rs("M_SEX")) <> "") then
%>			
						<tr>
							<td bgColor="<% =strPopUpTableColor %>" align=right nowrap valign=top><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">性&nbsp;&nbsp;&nbsp;&nbsp;别：&nbsp;</font></b></td>
							<td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% = ChkString(rs("M_SEX"), "display") %></font></td>
						</tr>
<%
				end if
				if (strOccupation = "1" and Trim(rs("M_OCCUPATION")) <> "") then
%>						<tr>
							<td bgColor="<% =strPopUpTableColor %>" align=right nowrap valign=top><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">现任职业：&nbsp;</font></b></td>
							<td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% = ChkString(rs("M_OCCUPATION"), "dislay") %></font></td>
						</tr>
<%
				end if
%>			
						<tr>
							<td bgColor="<% =strPopUpTableColor %>" align=right nowrap valign=top><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">文章总数：</font></b></td>
							<td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% = ChkString(rs("M_POSTS"), "display") %></font></td>
						</tr>
					
						
<%
			if not(strUseExtendedProfile)  then
%>
			<!--	 <tr>
					<td align=center bgcolor="<% =strCategoryCellColor %>" colspan="2"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">&nbsp;联络资讯&nbsp;</font></b></td>
				 </tr> -->
				 
				 <tr>
					<td bgColor=<% =strPopUpTableColor %> align=right width="10%" nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"> 电子邮件：&nbsp;</font></b></td>
					<td bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="pop_mail.asp?id=<% =rs("MEMBER_ID") %>"><% =ChkString(rs("M_EMAIL"), "display") %></a>&nbsp;</font></td>
				</tr>
<%				if strICQ = "1" then
					if Trim(rs("M_ICQ")) <> "" then %>
				<tr>
					<td bgColor=<% =strPopUpTableColor %> align=right nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">ICQ号码：&nbsp;</font></b></td>
					<td bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="pop_messengers.asp?mode=ICQ&ICQ=<% =ChkString(rs("M_ICQ"), "urlpath") %>&M_NAME=<% =ChkString(rs("M_NAME"), "urlpath") %>"><img src="http://online.mirabilis.com/scripts/online.dll?icq=<% =ChkString(rs("M_ICQ"), "urlpath") %>&img=5" border="0" align="absmiddle"><% =ChkString(rs("M_ICQ"), "display") %></a>&nbsp;</font></td>
				</tr>
<%					end if
				end if
				if strYAHOO = "1" then
					if Trim(rs("M_YAHOO")) <> "" then %>
				<tr>
					<td bgColor=<% =strPopUpTableColor %> align=right nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">OICQ号码：&nbsp;</font></b></td>
					<td bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="http://search.tencent.com/cgi-bin/friend/user_show_info?ln=<% =rs("M_YAHOO") %>"  target=_blank><img align=absmiddle width=16 height=16 src="http://icon.tencent.com/<% =rs("M_YAHOO") %>/s/00/" border=0><% =rs("M_YAHOO") %></a></font></td>
				</tr>
<%					end if
				end if
				if strAIM = "1" then
					if Trim(rs("M_AIM")) <> "" then %>
				<tr>
					<td bgColor=<% =strPopUpTableColor %> align=right nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">AIM号码：&nbsp;</font></b></td>
					<td bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="pop_messengers.asp?mode=AIM&AIM=<% =ChkString(rs("M_AIM"), "urlpath") %>&M_NAME=<% =ChkString(rs("M_NAME"), "urlpath") %>"><% =ChkString(rs("M_AIM"), "display") %></a>&nbsp;</font></td>
				</tr>
<%					end if
				end if 

			end if
%>
				
<%			if (strBio + strHobbies + strLNews + strQuote)	> 0 then %>			
				<tr>
					<td align=center bgcolor="<% =strCategoryCellColor %>" colspan="2"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">更多详细资料</font></b></td>
				</tr>
<%				if strBio = "1" then  %>
				<tr>				
					<td bgColor=<% =strPopUpTableColor %> valign=top align=right nowrap width="10%">
					<b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">个人简介：&nbsp;</font></b>
					</td>	
					<td bgColor=<% =strPopUpTableColor %> valign=top>
					<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% if IsNull(strMybio) or trim(strMyBio) = "" then Response.Write("-") else Response.Write(formatStr(strMyBio)) %></font>
					</td>
				</tr>
<%				end if
				if strHobbies = "1" then  %>
				<tr>
					<td bgColor=<% =strPopUpTableColor %> valign=top align=right nowrap width="10%">
					<b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">个人爱好：&nbsp;</font></b>
					</td>
					<td bgColor=<% =strPopUpTableColor %> >
					<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% if IsNull(strMyHobbies)  or trim(strMyHobbies) = "" then Response.Write("-") else Response.Write(formatStr(strMyHobbies)) %></font>
					</td>
				</tr>
<%				end if
				if strLNews = "1" then  %>
				<tr>
					<td bgColor=<% =strPopUpTableColor %> valign=top align=right nowrap width="10%">
					<b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">最近状况：&nbsp;</font></b>
					</td>
					<td bgColor=<% =strPopUpTableColor %> valign=top>
					<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% if IsNull(strMyLNews) or trim(strMyLNews) = "" then Response.Write("-") else Response.Write(formatStr(strMyLNews)) %></font>
					</td>
				</tr>
<%				end if
				if strQuote = "1" then  %>
				<tr>
					<td align=right bgcolor="<% =strPopUpTableColor %>" nowrap width="10%" valign=top>
					<b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">座 右 铭：&nbsp;</font></b></td>
					<td bgColor=<% =strPopUpTableColor %> valign=top>
					<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% if IsNull(strMyQuote) or Trim(strMyQuote) = "" then Response.Write("-") else Response.Write(formatStr(strMyQuote)) %></font></td>
				</tr>
<%				end if
			end if
'Rem User Field Code #######################################
			if (intUserFields	> 0) then %>
				<tr>
					<td align=center bgcolor="<% =strCategoryCellColor %>" colspan="2"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">更多详细资料</font></b></td>
				</tr>
				<tr><td></td></tr>
<%				set Conn = Server.CreateObject("ADODB.Connection")
				Conn.Open strConnString
				set rs2 = Conn.Execute("SELECT * FROM " & strMemberTablePrefix & "USERFIELDS ORDER BY USR_FIELD_ID")
				do until rs2.EOF
					Response.Write ("<tr><td valign=""top""  width=150 align=""right"" bgColor=" &strPopUpTableColor & "> <b><font face=""" & strDefaultFontFace & """ size=" & strDefaultFontSize & ">" & rs2("USR_LABEL") & ":</font></b></td>")	 
					if rs2("USR_FIELDTYPE") = "C" then
%>
					<td valign="top" bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% if getUserFieldValue(rs2("USR_FIELD_ID"),RS("MEMBER_ID")) = "1" then %>是<% Else  %>否<% End If %></font></td>			
					<% Else  %>
					<td valign="top" bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><%= getUserFieldValue(rs2("USR_FIELD_ID"),RS("MEMBER_ID")) %></font></td>			
					<% End If %>
					</tr>
<%					rs2.MoveNext
				loop
				rs2.Close
			end if
'Rem User Field Code #######################################

			if (strHomepage + strFavLinks) > 0 and not(strRecentTopics = "0" and strUseExtendedProfile) then  
				if strUseExtendedProfile then	
%>
				<tr>
					<td align=center bgcolor="<% =strCategoryCellColor %>" colspan="2">
					<b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">相关连接&nbsp;</font></b></td>
				</tr> 
<%			end if
				if strHomepage = "1" then %>
				<tr>
					<td bgColor=<% =strPopUpTableColor %> align=right nowrap width="10%"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">个人主页：&nbsp;</font></b></td>
<%					if Trim(rs("M_HOMEPAGE")) <> "" and lcase(trim(rs("M_HOMEPAGE"))) <> "http://" and Trim(lcase(rs("M_HOMEPAGE"))) <> "https://" then %>
					<td bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="<% = ChkString(rs("M_Homepage"), "urlpath") %>" target="_Blank"><% =ChkString(rs("M_Homepage"), "display") %></a>&nbsp;</font></td>
<%					else %>
					<td bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">没有个人主页</font></td>
<%					end if%>
				</tr>
<%				end if
				if strFavLinks = "1" then
					if Trim(rs("M_LINK1")) <> "" and lcase(trim(rs("M_LINK1"))) <> "http://" and Trim(lcase(rs("M_LINK1"))) <> "https://" then %>
				<tr>
					<td bgColor=<% =strPopUpTableColor %> align=right nowrap width="10%"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">推荐连接：&nbsp;</font></b></td>
					<td bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="<% =ChkString(rs("M_LINK1"), "urlpath") %>" target="_Blank"><% =ChkString(rs("M_LINK1"), "display") %></a>&nbsp;</font></td>
				</tr>
<%						if Trim(rs("M_LINK2")) <> "" and lcase(trim(rs("M_LINK2"))) <> "http://" and Trim(lcase(rs("M_LINK2"))) <> "https://" then %>
				<tr>
					<td bgColor=<% =strPopUpTableColor %> align=right nowrap width="10%"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">&nbsp;</font></b></td>
					<td bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="<% =ChkString(rs("M_LINK2"), "urlpath") %>" target="_Blank"><% =ChkString(rs("M_LINK2"), "display") %></a>&nbsp;</font></td>
				</tr>		
<%						end if
					end if
				end if
			end if
%>		
			</table>		
		 </td>
	    </tr>
	</table>		
</td>
</tr>
	</table>		
</td>
</tr>

	
<% if strUseExtendedProfile then %>	  
</table>
<table border="0" width="100%" cellspacing="0" cellpadding="3" valign=top>
	<tr>
		<td bgColor="<% =strPageBGColor %>" align=center nowrap>
			<br>
			<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript: onClick= history.go(-1) ">返回上一页</a></font></p>
			<p>&nbsp;</p>
		</td>
	  </tr>
	<tr>
<% else%>
	<tr>
	  <td bgColor="<% =strPageBGColor %>" align=center nowrap>
<% end if%>



<% case "Edit" %>
	<center>
	<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">用户资料</font></p>
	<p align=center><form action="pop_profile.asp?mode=goEdit&id=<% =Request.QueryString("id")%>" method="post">
	<input name="Refer" type="hidden" value="<% =Request.ServerVariables("HTTP_REFERER") %>">
	<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">修改你在论坛的注册资料.<br>
<%		if strAuthType = "nt" then %>
你的 NT 使用者帐号已显示。点选『送出』继续。<br><br>
<%		else %>
<%			if strAuthType = "db" then %>
请填写表格。<br><br>
<%			end if %>	
<%		end if %>	
	如果你尚未注册，<a href="policy.asp">按这里注册</a>。</center>

	<table border="0" cellspacing="0" cellpadding="0" align=center>
		<tr>
			<td bgcolor="<% =strPopUpBorderColor %>">
	<table border="0" width="100%" cellspacing="1" cellpadding="4">
<%		if strAuthType = "nt" then %>
		<TR>
			<TD bgColor="<% =strPopUpTableColor %>" align=right nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">你的帐号</font></b></td>
			<TD bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><% =Session(strCookieURL & "userid") %></font></b></td>
		</TR>
<%		else %>	
<%			if strAuthType = "db" then %>
	 	<TR>
	 	    <TD  bgColor=<% =strPopUpTableColor %> align=right nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">会员名称：</font></td>
	 	    <TD  bgColor=<% =strPopUpTableColor %>><input name="Name" size="25" value="<% =Request.Cookies(strUniqueID & "User")("Name")%>"></td>
	 	</TR>
	 	<TR>
	 	    <TD  bgColor=<% =strPopUpTableColor %> align=right nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">你的密码：</font></td>
	 	    <TD  bgColor=<% =strPopUpTableColor %>><input name="Password" type=Password size="25" value="<% =Request.Cookies(strUniqueID & "User")("Pword")%>"></td>
	 	</TR>
<%			end if %>
<%		end if %>
	 	<TR>
	 		<TD  bgColor=<% =strPopUpTableColor %> align=center colspan=2><input type=submit value=提交></td>
	 	</TR>    
	</table>
			</td>
		</tr>
	</table>
	
	</form></p>
<%
	case "goEdit"

		if strAuthType = "db" then
			if strDBNTUserName = "" then 
				strDBNTUserName = Request.Form("Name")
			end if
		end if


		'## Forum_SQL
		strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME " 
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_USERNAME, " & strMemberTablePrefix & "MEMBERS.M_EMAIL "
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_FIRSTNAME, " & strMemberTablePrefix & "MEMBERS.M_LASTNAME " 
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_LEVEL"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_TITLE"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_PASSWORD"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_ICQ"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_YAHOO"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_AIM"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_COUNTRY "
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_POSTS"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_CITY " 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_STATE " 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_COUNTRY " 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_POSTS " 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_HIDE_EMAIL " 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_DATE " 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_PHOTO_URL " 
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_HOMEPAGE" 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_LINK1, " & strMemberTablePrefix & "MEMBERS.M_LINK2 "
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_AGE, " & strMemberTablePrefix & "MEMBERS.M_MARSTATUS " 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_SEX, " & strMemberTablePrefix & "MEMBERS.M_OCCUPATION " 
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_SIG"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_HOBBIES, " & strMemberTablePrefix & "MEMBERS.M_LNEWS " 
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_QUOTE, " & strMemberTablePrefix & "MEMBERS.M_BIO "
'######################################## kerrycode for Mail List # Display in Edit mode
		
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_RECMAIL"

'-------/kerrycode		
		strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " WHERE "&Strdbntsqlname&" = '" & ChkString(STRdbntUserName, "decode") & "' "
			if strAuthType = "db" then
				strSql = strSql & " AND   M_PASSWORD = '" & ChkString(Request.Form("Password"),"password") & "'"
			end if
		QuoteOk = ChkQuoteOk(STRdbntUserName)
		set rs = my_Conn.Execute(strSql)
'Response.write strSQL 
'response.end
		if rs.BOF and rs.EOF and QuoteOk then 
%>
	<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">错误的用户名或密码</font></p>
	
	<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript:history.go(-1)">返回重试</a></font></p>
<%
			if strUseExtendedProfile then 
%>
	<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="<% if InStr(1,Request.Form("refer"),"register.asp",1) > 0 then Response.Write("default.asp") else Response.Write(Request.Form("refer")) end if %>">返回论坛</a></font></p>
<%
			end if 
		else
			'## Display Edit Profile Page
%>
	<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">编辑注册资料</font></p>
	<p align=center><form action="pop_profile.asp?mode=EditIt&id=<% =Request.QueryString("id")%>" method="Post" id=Form1 name=Form1>
	<input name="Refer" type="hidden" value="<% =Request.Form("Refer") %>">
	<!-- #include file="inc_profile.asp" -->
	</form></p>
<%
		end if 

	case "Modify"
%>
<center>
<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">编辑会员资料</font></p>

<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b><font color=red>注意：</font></b>只有管理员才能修改会员资料</font></p>
<center>
<form action="pop_profile.asp?mode=goModify" method=post id=Form1 name=Form1>
<input type=hidden name="Method_Type" value="<% =Request.QueryString("mode") %>">
<input type=hidden name="MEMBER_ID" value="<% =Request.QueryString("ID") %>">
<input type=hidden name="Name" value="<% =Request.QueryString("name") %>">
<input name="Refer" type="hidden" value="<% =Request.ServerVariables("HTTP_REFERER") %>">

<table border="0" width="30%" cellspacing="0" cellpadding="0" align=center>
  <tr>
    <td bgcolor="<% =strPopUpBorderColor %>">
    <table border="0" width="100%" cellspacing="1" cellpadding="4">
<% if strAuthType="db" then %>
      <tr>
        <TD bgColor=<% =strPopUpTableColor %> align=right nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">用户名称：</font></b></td>
        <TD bgColor=<% =strPopUpTableColor %>><input type=text name="User" value="<% =Request.Cookies(strUniqueID & "User")("Name")%>" size=20></td>
      </tr>
      <tr>
        <TD bgColor=<% =strPopUpTableColor %> align=right nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">你的密码：</FONT></b></td>
        <TD bgColor=<% =strPopUpTableColor %>><input type=Password name="Pass" value="<% =Request.Cookies(strUniqueID & "User")("Pword")%>" size=20></td>
      </tr>
<% elseif strAuthType="nt" then %>
	<tr>
	  <td bgcolor=<% =strPopUpTableColor %> align=right nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">NT 帐号：</font></b></td>
	  <td bgcolor=<% =strPopUpTableColor %>><font face="<%=strDefaultFontFace %>" size="<% =strDefaultFontSize%>"><%=Session(strCookieURL & "userid")%></FONT></td>
	</tr>
<% end if %>
      <tr>
        <TD bgColor=<% =strPopUpTableColor %> colspan=2 align=center><Input type=Submit value="提交" id=Submit1 name=Submit1></td>
      </TR>
    </TABLE>
    </td>
  </tr>
</table>
</form>
<%
	case "goModify"

		mLev = cint(ChkUser2(STRdbntUserName, Request.Form("Pass"))) 
						
		if mLev > 0 then  '## is Member
			if mLev = 4 then

			'## Forum_SQL
			strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME " 
			strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_USERNAME, " & strMemberTablePrefix & "MEMBERS.M_EMAIL "
			strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_FIRSTNAME, " & strMemberTablePrefix & "MEMBERS.M_LASTNAME " 
			strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_LEVEL"
			strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_TITLE"
			strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_PASSWORD"
			strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_ICQ"
			strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_YAHOO"
			strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_AIM"
			strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_COUNTRY "
			strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_POSTS"
			strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_CITY " 
			strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_STATE " 
			strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_COUNTRY " 
			strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_POSTS " 
			strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_HIDE_EMAIL " 
			strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_DATE " 
			strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_PHOTO_URL " 
			strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_HOMEPAGE" 
			strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_LINK1, " & strMemberTablePrefix & "MEMBERS.M_LINK2 "
			strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_AGE, " & strMemberTablePrefix & "MEMBERS.M_MARSTATUS " 
			strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_SEX, " & strMemberTablePrefix & "MEMBERS.M_OCCUPATION " 
			strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_SIG"
			strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_HOBBIES, " & strMemberTablePrefix & "MEMBERS.M_LNEWS " 
			strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_QUOTE, " & strMemberTablePrefix & "MEMBERS.M_BIO "
			strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS.M_ALLOWDOWNLOADS, " & strMemberTablePrefix & "MEMBERS.M_ALLOWUPLOADS "
'######################################## kerrycode for Mail List # Display in goModify mode
		
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS.M_RECMAIL"

'-------/kerrycode			
			strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
			strSql = strSql & " WHERE MEMBER_ID = " & Request.Form("MEMBER_ID")

			
			set rs = my_Conn.Execute(strSql)

			'## Display Edit Profile Page
%>
	<center>
	<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">编辑会员资料</font></p>
	<p align=center><form action="pop_profile.asp?mode=ModifyIt&id=<% =Request.Form("MEMBER_ID")%>" method="Post" id=Form1 name=Form1>
	</center>
	<input type=hidden name="User" value="<% =Request.Form("User") %>">
	<input type=hidden name="Pass" value="<% =Request.Form("Pass") %>">
	<input name="Refer" type="hidden" value="<% =Request.Form("Refer") %>">
	<!-- #include file="inc_profile.asp" -->
	</form></p>
<%			else %>
<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><b>你不是管理员不能编辑会员资料</b></font><br>
<br>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript: onClick= history.go(-1) ">返回重新认证</a></font></p>
<%				if strUseExtendedProfile then %>
	<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="<% if InStr(1,Request.Form("refer"),"register.asp",1) > 0 then Response.Write("default.asp") else Response.Write(Request.Form("refer")) end if %>">返回论坛</a></font></p>

<%				end if 
			end if 
		else %>
<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><b>你不是管理员不能编辑会员资料</b></font><br>
<br>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript: onClick= history.go(-1) ">返回重新认证</a></font></p>

<%			if strUseExtendedProfile then %>
	<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="<% if InStr(1,Request.Form("refer"),"register.asp",1) > 0 then Response.Write("default.asp") else Response.Write(Request.Form("refer")) end if %>">返回论坛</a></font></p>

<%			end if 
		end if 

	case "EditIt"

		Err_Msg = ""
		if Request.Form("Name") = "" then 
			Err_Msg = Err_Msg & "<li>你必须输入姓名</li>"
		end if
		if (Instr(Request.Form("Name"), ">") > 0 ) or (Instr(Request.Form("Name"), "<") > 0) then
			Err_Msg = Err_Msg & "<li> &gt; 和 &lt; 不能出现在用户名里，请重新输入</li>"
		end if
		if strAuthType = "db" then
			if Request.Form("Password") = "" then 
				Err_Msg = Err_Msg &  "<li>你必须输入密码</li>"
			end if
			if Len(Request.Form("Password")) > 25 then 
				Err_Msg = Err_Msg & "<li>密码不能超过25个字元</li>" 
			end if
			if Request.Form("Password") <>  Request.Form("Password2") then 
				Err_Msg = Err_Msg & "<li>两次输入的密码不同</li>"
			end if
		end if
		if Request.Form("Email") = "" then 
			Err_Msg = Err_Msg & "<li>必须填写电子邮件</li>"
		end if
		if EmailField(Request.Form("Email")) = 0 then 
			Err_Msg = Err_Msg & "<li>必须填写正确的电子邮件</li>"
		end if
		if (lcase(left(Request.Form("Homepage"), 7)) <> "http://") and (lcase(left(Request.Form("Homepage"), 8)) <> "https://") and (Request.Form("Homepage") <> "") then
			Err_Msg = Err_Msg & "<li>你必须在网址前加上 <b>http://</b> or <b>https://</b></li>"
		end if
		if strUniqueEmail = "1" then
			if lcase(Request.Form("Email")) <> lcase(Request.Form("Email2")) then
				'## Forum_SQL
				strSql = "SELECT M_EMAIL FROM " & strMemberTablePrefix & "MEMBERS "
				strSql = strSql & " WHERE M_EMAIL = '" & Trim(Request.Form("Email")) &"'"

				set rs = my_Conn.Execute (strSql)

				if rs.BOF and rs.EOF then 
					'## Do Nothing - proceed
				Else 
					Err_Msg = Err_Msg & "<li>此电子邮件已经有人使用，请重新输入</li>"
				end if
				rs.close
				set rs = nothing
			end if
		end if
		if Len(ChkString(Request.Form("Sig"),"message")) > 255 then
				Err_Msg = Err_Msg & "<li>个人签名不能超过255个字符。"
				Err_Msg = Err_Msg & "现在共有 <b>" & Len(ChkString(Request.Form("Sig"),"message")) & "</b> 个字符</li>"
		end if
		if Err_Msg = "" then
		
			if Trim(Request.Form("Homepage")) <> "" and lcase(trim(Request.Form("Homepage"))) <> "http://" and Trim(lcase(Request.Form("Homepage"))) <> "https://" then
				regHomepage = ChkString(Request.Form("Homepage"),"url")
			else
				regHomepage = " "
			end if
			if Trim(Request.Form("LINK1")) <> "" and lcase(trim(Request.Form("LINK1"))) <> "http://" and Trim(lcase(Request.Form("LINK1"))) <> "https://" then
				regLink1 = ChkString(Request.Form("LINK1"),"url")
			else
				regLink1 = " "
			end if
			if Trim(Request.Form("LINK2")) <> "" and lcase(trim(Request.Form("LINK2"))) <> "http://" and Trim(lcase(Request.Form("LINK2"))) <> "https://" then
				regLink2 = ChkString(Request.Form("LINK2"),"url")
			else
				regLink2 = " "
			end if
			if Trim(Request.Form("PHOTO_URL")) <> "" and lcase(trim(Request.Form("PHOTO_URL"))) <> "http://" and Trim(lcase(Request.Form("PHOTO_URL"))) <> "https://" then
				regPhoto_URL = ChkString(Request.Form("Photo_URL"),"url")
			else
				regPhoto_URL = " "
			end if

			'## Forum_SQL
			strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
			strSql = strSql & " SET M_PASSWORD = '" & ChkString(Request.Form("Password"),"") & "', "
			strSql = strSql & "     M_COUNTRY  = '" & ChkString(Request.Form("Country"),"")  & "', "
'#########################kerrycode for Email List # UPDATE in Edit
			
			strSql = strSql & "     M_RECMAIL  = '" & ChkString(Request.Form("RECMAIL"),"")  & "', "
			
'----------------------/kerrycode			
			if strICQ = "1" then
				strSql = strSql & "     M_ICQ = '" & ChkString(Request.Form("ICQ"),"") & "', "
			end if
			if strYAHOO = "1" then
				strSql = strSql & "     M_YAHOO = '" & ChkString(Request.Form("YAHOO"),"") & "', "
			end if
			if strAIM = "1" then
				strSql = strSql & "     M_AIM = '" & ChkString(Request.Form("AIM"),"") & "', "
			end if
			if strHOMEPAGE = "1" then
				strSql = strSql & "     M_Homepage = '" & ChkString(Trim(regHomepage),"") & "', "
			end if
			strSql = strSql & "     M_SIG = '" & ChkString(Request.Form("Sig"),"message") & "', "
			strSql = strSql & "     M_EMAIL = '" & ChkString(Request.Form("Email"),"") & "' "
			if strfullName = "1" then
				strSql = strSql & ",	M_FIRSTNAME = '" & ChkString(Request.Form("FirstName"),"") & "'" 
				strSql = strSql & ",	M_LASTNAME  = '" & ChkString(Request.Form("LastName"),"") & "'"  
			end if
			if strCity = "1" then
				strsql = strsql & ",	M_CITY = '" & ChkString(Request.Form("City"),"") & "'"  
			end if
			if strState = "1" then
				strsql = strsql & ",	M_STATE = '" & ChkString(Request.Form("State"),"") & "'" 
			end if
			strsql = strsql & ",	M_HIDE_EMAIL = '" & ChkString(Request.Form("HideMail"),"") & "'"  
			if strPicture = "1" then
				strsql = strsql & ",	M_PHOTO_URL = '" & ChkString(Request.Form("Photo_URL"),"") & "'"  
			end if
			if strFavLinks = "1" then
				strsql = strsql & ",	M_LINK1 = '" & ChkString(Request.Form("LINK1"),"") & "'" 
				strSql = strSql & ",	M_LINK2 = '" & ChkString(Request.Form("LINK2"),"") & "'" 
			end if
			if strAge = "1" then
				strSql = strsql & ",	M_AGE = '" & ChkString(Request.Form("Age"),"") & "'" 
			end if
			if strMarStatus = "1" then
				strSql = strSql & ",	M_MARSTATUS = '" & ChkString(Request.Form("MarStatus"),"") & "'" 
			end if
			if strSex = "1" then
				strSql = strsql & ",	M_SEX = '" & ChkString(Request.Form("Sex"),"") & "'" 
			end if
			if strOccupation = "1" then
				strSql = strSql & ",	M_OCCUPATION = '" & ChkString(Request.Form("Occupation"),"") & "'" 
			end if
			if strBio = "1" then
				strSql = strSql & ",	M_BIO = '" & ChkString(Request.Form("Bio"),"message") & "'" 
			end if
			if strHobbies = "1" then
				strSql = strSql & ",	M_HOBBIES = '" & ChkString(Request.Form("Hobbies"),"message") & "'" 
			end if
			if strLNews = "1" then
				strsql = strsql & ",	M_LNEWS = '" & ChkString(Request.Form("LNews"),"message") & "'" 
			end if
			if strQuote = "1" then
				strSql = strSql & ",	M_QUOTE = '" & ChkString(Request.Form("Quote"),"message") & "'" 
			end if
			strSql = strSql & " WHERE M_NAME = '" & Request.Form("Name") & "' "
			if strAuthType = "db" then 
			strSql = strSql & " AND   M_PASSWORD = '" & Request.Form("Password-d") & "'"
			end IF			

			my_Conn.Execute(strSql)
			
'Rem User Field Code #######################################
			if (strUseExtendedProfile) then
				if (intUserFields > 0) then
				
					set conn = Server.CreateObject("ADODB.Connection")
					
					conn.open strConnString

					memID = getMemberID(Request.Form("Name"))
					UserstrSql = "SELECT * " 
					UserstrSql = UserstrSql & " FROM " & strMemberTablePrefix & "USERFIELDS "
					set rsIDs = Conn.Execute(UserstrSql)
					do until rsIDs.EOF
						sInputName = rsIDs("USR_SHORTNAME")
						if DoesUserFieldExist(memID,rsIDs("USR_FIELD_ID")) then
							UpdSQL = "UPDATE " & strMemberTablePrefix & "MEMBERFIELDS "
							UpdSQL = UpdSQL & "SET " & strMemberTablePrefix & "MEMBERFIELDS.USR_VALUE ='" & Request.Form(sInputName) & "'"
							UpdSQL = UpdSQL & "WHERE " & strMemberTablePrefix & "MEMBERFIELDS.USR_FIELD_ID =" & rsIDs("USR_FIELD_ID") & " AND "
							UpdSQL = UpdSQL &  strMemberTablePrefix & "MEMBERFIELDS.MEMBER_ID =" & memID
							Conn.Execute(UpdSQL)
							UpdSQL = ""
						else
							UpdSql = "INSERT INTO " & strMemberTablePrefix & "MEMBERFIELDS "
							UpdSql = UpdSql & "VALUES ( " & memID & "," & rsIDs("USR_FIELD_ID") & ",'"
							UpdSql = UpdSql & Request.Form(sInputName) & "')"
							Conn.Execute(UpdSQL)
							UpdSQL = ""
						end if
						rsIDs.MoveNext
					loop
					rsIDs.Close
					Err_Msg = "OK"
				end if
			end if
'Rem User Field Code #######################################
			regHomepage = ""
			
%>
	<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">你的注册资料已经修改完毕！</font></p>
<%
					if (strUseExtendedProfile) then
%>					
					<meta http-equiv="Refresh" content="0; URL=<% if InStr(1,Request.Form("refer"),"register.asp",1) > 0 then Response.Write("default.asp") else Response.Write(Request.Form("refer")) end if %>">
					<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="<% if InStr(1,Request.Form("refer"),"register.asp",1) > 0 then Response.Write("default.asp") else Response.Write(Request.Form("refer")) end if %>">返回论坛</a></font></p>
<%
					end if
			else

%>
	<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">你输入的资料有问题或没有填写完整</font></p>

	<table align=center>
	  <tr>
	    <td align=center><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><center>
		<ul>
		<% =Err_Msg %>
		</ul></center></font>
	    </td>
	  </tr>
	</table>

	<p align=center><font size="<% =strDefaultFontSize %>"><a href="JavaScript:history.go(-1)">请返回重新输入</a></font></p>
<%			if strUseExtendedProfile then %>
	<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="<% if InStr(1,Request.Form("refer"),"register.asp",1) > 0 then Response.Write("default.asp") else Response.Write(Request.Form("refer")) end if %>">返回论坛</a></font></p>

<%			end if 
		end if

	case "ModifyIt"

		mLev = cint(ChkUser2(STRdbntUserName, Request.Form("Pass"))) 
						
		if mLev > 0 then  '## is Member
			if mLev = 4 then '## is Admin

				Err_Msg = ""
			
				if Request.Form("Name") = "" then 
					Err_Msg = Err_Msg & "<li>You must set a UserName</li>"
				end if
				if (Instr(Request.Form("Name"), ">") > 0 ) or (Instr(Request.Form("Name"), "<") > 0) then
					Err_Msg = Err_Msg & "<li> &gt; 和 &lt; 不能出现在用户名里，请重新输入</li>"
				end if
				'## Forum_SQL
				strSql = "SELECT M_NAME FROM " & strMemberTablePrefix & "MEMBERS "
				strSql = strSql & " WHERE M_NAME = '" & Trim(Request.Form("Name")) &"' "
				strSql = strSql & " AND MEMBER_ID <> " & Trim(Request.Form("Member_ID")) &" "
		
				set rs = my_Conn.Execute (strSql)	

				if rs.BOF and rs.EOF then 
					'## Do Nothing - proceed
				else 
					Err_Msg = Err_Msg & "<li>用户名重复<br>请重新输入</li>"
				end if
						
				rs.close
				set rs = nothing
				if strAuthType = "db" then
					if Request.Form("Password") = "" then 
						Err_Msg = Err_Msg &  "<li>你必须设定密码</li>"
					end if
					if Len(Request.Form("Password")) > 25 then 
						Err_Msg = Err_Msg & "<li>密码不能超过25个字符</li>" 
					end if
					'if Request.Form("Password") <>  Request.Form("Password2") then 
					'	Err_Msg = Err_Msg & "<li>两次输入密码不一致</li>"
					'end if
				end if
				if Request.Form("Email") = "" then 
					Err_Msg = Err_Msg & "<li>你必须填写电子邮件</li>"
				end if
				if EmailField(Request.Form("Email")) = 0 then 
					Err_Msg = Err_Msg & "<li>必须填写正确的电子邮件地址</li>"
				end if
				if (lcase(left(Request.Form("Homepage"), 7)) <> "http://") and (lcase(left(Request.Form("Homepage"), 8)) <> "https://") and (Request.Form("Homepage") <> "") then
					Err_Msg = Err_Msg & "<li>你必须在连接前加上 <b>http://</b> or <b>https://</b></li>"
				end if
				if Len(Request.Form("Sig")) > 255 then
					Err_Msg = Err_Msg & "<li>个人签名不能超过255个字符 "
					Err_Msg = Err_Msg & "现在共有 <b>" & Len(Request.Form("Sig")) & "</b> 个字符</li>"
				end if
				if Err_Msg = "" then '## it is ok to update the profile
		
					if Trim(Request.Form("Homepage")) <> "" and lcase(trim(Request.Form("Homepage"))) <> "http://" and Trim(lcase(Request.Form("Homepage"))) <> "https://" then
						regHomepage = ChkString(Request.Form("Homepage"),"url")
					else
						regHomepage = " "
					end if
					if Trim(Request.Form("LINK1")) <> "" and lcase(trim(Request.Form("LINK1"))) <> "http://" and Trim(lcase(Request.Form("LINK1"))) <> "https://" then
						regLink1 = ChkString(Request.Form("LINK1"),"url")
					else
						regLink1 = " "
					end if
					if Trim(Request.Form("LINK2")) <> "" and lcase(trim(Request.Form("LINK2"))) <> "http://" and Trim(lcase(Request.Form("LINK2"))) <> "https://" then
						regLink2 = ChkString(Request.Form("LINK2"),"url")
					else
						regLink2 = " "
					end if
					if Trim(Request.Form("PHOTO_URL")) <> "" and lcase(trim(Request.Form("PHOTO_URL"))) <> "http://" and Trim(lcase(Request.Form("PHOTO_URL"))) <> "https://" then
						regPhoto_URL = ChkString(Request.Form("Photo_URL"),"url")
					else
						regPhoto_URL = " "
					end if
			
					'## Forum_SQL
					strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
					strSql = strSql & " SET M_NAME = '" & ChkString(Request.Form("Name"),"name") & "'"
					if strAuthType = "nt" then
						strSql = strSql & ",    M_USERNAME = '" & ChkString(Request.Form("Account"),"") & "'"
					else
						if strAuthType = "db" then
							strSql = strSql & ",    M_PASSWORD = '" & ChkString(Request.Form("Password"),"") & "'"
						end if
					end if
					strSql = strSql & ",    M_EMAIL = '" & ChkString(Request.Form("Email"),"") & "'"
'#########################kerrycode for Email List # UPDATE in Modify
			
			strSql = strSql & ",     M_RECMAIL  = '" & ChkString(Request.Form("RECMAIL"),"")  & "'"
			
'----------------------/kerrycode					
					strSql = strSql & ",    M_TITLE = '" & ChkString(Request.Form("Title"),"") & "'"
					strSql = strSql & ",    M_POSTS = " & ChkString(Request.Form("Posts"),"") & " "
					strSql = strSql & ",    M_COUNTRY = '" & ChkString(Request.Form("Country"),"") & "'"
					if strICQ = "1" then
						strSql = strSql & ",    M_ICQ = '" & ChkString(Request.Form("ICQ"),"") & "'"
					end if
					if strYAHOO = "1" then
						strSql = strSql & ",    M_YAHOO = '" & ChkString(Request.Form("YAHOO"),"") & "'"
					end if
					if strAIM = "1" then
						strSql = strSql & ",    M_AIM = '" & ChkString(Request.Form("AIM"),"name") & "'"
					end if
					if strHOMEPAGE = "1" then
						strSql = strSql & ",    M_HOMEPAGE = '" & ChkString(Request.Form("Homepage"),"" ) & "'"
					end if
					strSql = strSql & ",    M_SIG = '" & ChkString(Request.Form("Sig"),"message") & "'"
					strSql = strSql & ",    M_LEVEL = " & ChkString(Request.Form("Level"),"")
					if strfullName = "1" then
						strSql = strSql & ",	M_FIRSTNAME = '" & ChkString(Request.Form("FirstName"),"") & "'" 
						strSql = strSql & ",	M_LASTNAME  = '" & ChkString(Request.Form("LastName"),"") & "'"  
					end if
					if strCity = "1" then
						strsql = strsql & ",	M_CITY = '" & ChkString(Request.Form("City"),"") & "'"  
					end if
					if strState = "1" then
						strsql = strsql & ",	M_STATE = '" & ChkString(Request.Form("State"),"") & "'" 
					end if
					strsql = strsql & ",	M_HIDE_EMAIL = '" & ChkString(Request.Form("HideMail"),"") & "'"  
					if strPicture = "1" then
						strsql = strsql & ",	M_PHOTO_URL = '" & ChkString(Request.Form("Photo_URL"),"") & "'"  
					end if
					if strFavLinks = "1" then
						strsql = strsql & ",	M_LINK1 = '" & ChkString(Request.Form("LINK1"),"") & "'" 
						strSql = strSql & ",	M_LINK2 = '" & ChkString(Request.Form("LINK2"),"") & "'" 
					end if
					if strAge = "1" then
						strSql = strsql & ",	M_AGE = '" & ChkString(Request.Form("Age"),"") & "'" 
					end if
					if strMarStauts = "1" then
						strSql = strSql & ",	M_MARSTATUS = '" & ChkString(Request.Form("MarStatus"),"") & "'" 
					end if
					if strSex = "1" then
						strSql = strsql & ",	M_SEX = '" & ChkString(Request.Form("Sex"),"") & "'" 
					end if
					if strOccupation = "1" then
						strSql = strSql & ",	M_OCCUPATION = '" & ChkString(Request.Form("Occupation"),"") & "'" 
					end if
					if strBio = "1" then
						strSql = strSql & ",	M_BIO = '" & ChkString(Request.Form("Bio"),"message") & "'" 
					end if
					if strHobbies = "1" then
						strSql = strSql & ",	M_HOBBIES = '" & ChkString(Request.Form("Hobbies"),"message") & "'" 
					end if
					if strLNews = "1" then
						strsql = strsql & ",	M_LNEWS = '" & ChkString(Request.Form("LNews"),"message") & "'" 
					end if
					if strQuote = "1" then
						strSql = strSql & ",	M_QUOTE = '" & ChkString(Request.Form("Quote"),"message") & "'" 
					end if
'File Attachment ################################################
					if Request.Form("allowDownloads") = 1 then
						strSql = strSql & ",	M_ALLOWDOWNLOADS = 1"
					else 					
						strSql = strSql & ",	M_ALLOWDOWNLOADS = 0"
					end if
					if Request.Form("allowUploads") = 1 then
						strSql = strSql & ",	M_ALLOWUPLOADS = 1"
					else 					
						strSql = strSql & ",	M_ALLOWUPLOADS = 0"
					end if
'File Attachment ##################################################					
					strSql = strSql & " WHERE MEMBER_ID = " & Request.Form("MEMBER_ID")	

					my_Conn.Execute(strSql)
					
					if ChkString(Request.Form("Level"),"") = "1" then 
						'## Forum_SQL - Remove the member from the moderator table
						strSql = "DELETE FROM " & strTablePrefix & "MODERATOR "
						strSql = strSql & " WHERE " & strTablePrefix & "MODERATOR.MEMBER_ID = " & Request.Form("MEMBER_ID")

						my_Conn.Execute (strSql)
					end if

					
%>
					<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">你的注册资料已经修改完毕！</font></p>
<%
					if (strUseExtendedProfile) then
%>					
					<meta http-equiv="Refresh" content="0; URL=<% if InStr(1,Request.Form("refer"),"register.asp",1) > 0 then Response.Write("default.asp") else Response.Write(Request.Form("refer")) end if %>">
					<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="<% if InStr(1,Request.Form("refer"),"register.asp",1) > 0 then Response.Write("default.asp") else Response.Write(Request.Form("refer")) end if %>">返回论坛</a></font></p>
<%
					end if
				else

%>
					<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">输入资料有误或没有填写完整</font></p>
					<table align=center>
					  <tr>
						<td align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><center>
							<ul>
								<% =Err_Msg %>
							</ul>
							</center></font>
						</td>
					  </tr>
					</table>

					<p align=center><font size="<% =strDefaultFontSize %>"><a href="JavaScript:history.go(-1)">请返回重新输入</a></font></p>
<%					if strUseExtendedProfile then %>
					<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="<% if InStr(1,Request.Form("refer"),"register.asp",1) > 0 then Response.Write("default.asp") else Response.Write(Request.Form("refer")) end if %>">返回论坛</a></font></p>

<%					end if 
				end if
			else 'Member but no Admin%>
				<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><b>你不是管理员不能修改会员资料</b></font><br>
				<br>
				<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript: onClick= history.go(-1) ">返回重新认证</a></font></p>
<%				if strUseExtendedProfile then %>
					<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="<% if InStr(1,Request.Form("refer"),"register.asp",1) > 0 then Response.Write("default.asp") else Response.Write(Request.Form("refer")) end if %>">返回论坛</a></font></p>
<%				end if 
			end if %>
<%		else  'Not logged on or no member%>
			<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><b>你不是管理员不能修改会员资料</b></font><br>
			<br>
			<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript: onClick= history.go(-1) ">返回重新认证</a></font></p>
<%			if strUseExtendedProfile then %>
			<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="<% if InStr(1,Request.Form("refer"),"register.asp",1) > 0 then Response.Write("default.asp") else Response.Write(Request.Form("refer")) end if %>">返回论坛</a></font></p>
<%			end if 
		end if %>
	
<%

end select
on error resume next
rs.close
set rs=nothing
if not(strUseExtendedProfile) then
%>
	<!--#INCLUDE FILE="inc_footer_short.asp" -->
<%
else
%>
	<!--#INCLUDE FILE="inc_footer.asp" -->
<% end if %>
