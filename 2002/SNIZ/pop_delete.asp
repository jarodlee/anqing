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
<!--#INCLUDE FILE="inc_top_short.asp" -->
<% 
if strAuthType = "db" then
	STRdbntUserName = Request.Form("User")
end if

'################### File Attachments ############################ 
function DeleteUploads(msgType,ID)
	select case msgType
		case "Topic"
			strSql = "SELECT USERFILES.MEMBER_ID,USERFILES.F_FILENAME FROM " & strTablePrefix & "USERFILES "
			strSql = strSql & " WHERE " & strTablePrefix & "USERFILES.F_TOPIC_ID = " & ID
			set rs2 = my_Conn.Execute (strSql)
			on error resume next
			if not(rs2.eof or rs2.bof) then
				Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
				do while not rs2.eof
				     FilePath = replace(strCookieURL,"/mods/","/",1,-1,1) & "uploaded"
					 FilePath = Server.MapPath(FilePath) & "\" & getMemberName(rs2("MEMBER_ID")) & "\"
					if ScriptObject.FileExists(FilePath & rs2("F_FILENAME")) then
					ScriptObject.DeleteFile(FilePath & rs2("F_FILENAME"))
					end if
					rs2.MoveNext
				loop
				DeleteUploads = 1
			else 
				DeleteUploads = 0
			end if
			rs2.Close
		case "Reply"
			strSql = "SELECT MEMBER_ID,F_FILENAME FROM " & strTablePrefix & "USERFILES "
			strSql = strSql & " WHERE F_REPLY_ID = " & ID
			set rs2 = my_Conn.Execute (strSql)
			if not(rs2.eof or rs2.bof) then
				Set ScriptObject = Server.CreateObject("Scripting.FileSystemObject")
				on error resume next
				do while not rs2.eof
				     FilePath = replace(strCookieURL,"/mods/","/",1,-1,1) & "uploaded"
					 FilePath = Server.MapPath(FilePath) & "\" & getMemberName(rs2("MEMBER_ID")) & "\"
					 if ScriptObject.FileExists(FilePath & rs2("F_FILENAME")) then
						ScriptObject.DeleteFile(FilePath & rs2("F_FILENAME"))
					end if
					rs2.MoveNext
				loop
				DeleteUploads = 1
			else
				DeleteUploads = 0
			end if
			rs2.Close
	end Select
end function
'################### File Attachment end ############################

if Request.QueryString("mode") = "DeleteReply" then 
	mLev = cint(ChkUser3(strDBNTUserName, Request.Form("Pass"), Request.Form("REPLY_ID"))) 
	if mLev > 0 then  '## is Member
	  if (chkForumModerator(Request.Form("FORUM_ID"), strDBNTUserName) = "1") or (mLev = 1) or (mLev = 4) then '## is Allowed
			'## Forum_SQL - Delete reply
			strSql = "DELETE FROM " & strTablePrefix & "REPLY "
			strSql = strSql & " WHERE " & strTablePrefix & "REPLY.REPLY_ID = " & Request.Form("REPLY_ID")
			my_Conn.Execute strSql

'################### File Attachments ############################
			if DeleteUploads("Reply",Request.Form("REPLY_ID")) = 1 then
			strSql = "DELETE FROM " & strTablePrefix & "USERFILES "
			strSql = strSql & " WHERE " & strTablePrefix & "USERFILES.F_REPLY_ID = " & Request.Form("REPLY_ID")
			my_Conn.Execute strSql
			end if
'################### File Attachments end ########################
			
			set rs = Server.CreateObject("ADODB.Recordset")

			'## Forum_SQL - Get last_post and last_post_author for Topic
			strSql = "SELECT R_DATE, R_AUTHOR"
			strSql = strSql & " FROM " & strTablePrefix & "REPLY "
			strSql = strSql & " WHERE TOPIC_ID = " & Request.Form("TOPIC_ID") & " "
			strSql = strSql & " ORDER BY R_DATE DESC"

			set rs = my_Conn.Execute (strSql)
			
			if not(rs.eof or rs.bof) then
				strLast_Post = rs("R_DATE")
				strLast_Post_Author = rs("R_AUTHOR")
			end if			
			if (rs.eof or rs.bof) or IsNull(strLast_Post) or IsNull(strLast_Post_Author) then  'topic has no replies
				set rs2 = Server.CreateObject("ADODB.Recordset")

				'## Forum_SQL - Get post_date and author from Topic
				strSql = "SELECT T_AUTHOR, T_DATE "
				strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
				strSql = strSql & " WHERE TOPIC_ID = " & Request.Form("TOPIC_ID") & " "
				
				set rs2 = my_Conn.Execute (strSql)
			
				strLast_Post = rs2("T_DATE")
				strLast_Post_Author = rs2("T_AUTHOR")
				
				rs2.Close
				set rs2 = nothing
				
			end if
			
			rs.Close
			set rs = nothing
			
			'## FORUM_SQL - Decrease count of replies to individual topic by 1
			strSql = "UPDATE " & strTablePrefix & "TOPICS "
			strSql = strSql & " SET " & strTablePrefix & "TOPICS.T_REPLIES = " & strTablePrefix & "TOPICS.T_REPLIES - " & 1 & " "
			if strLast_Post <> "" then 
				strSql = strSql & ", T_LAST_POST = '" & strLast_Post & "'"
				if strLast_Post_Author <> "" then 
					strSql = strSql & ", T_LAST_POST_AUTHOR = " & strLast_Post_Author & ""
				end if
			end if
			strSql = strSql & " WHERE " & strTablePrefix & "TOPICS.TOPIC_ID = " & Request.Form("TOPIC_ID")

			my_Conn.Execute strSql

			'## Forum_SQL - Get last_post and last_post_author for Forum
			strSql = "SELECT T_LAST_POST, T_LAST_POST_AUTHOR "
			strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
			strSql = strSql & " WHERE FORUM_ID = " & Request.Form("FORUM_ID") & " "
			strSql = strSql & " ORDER BY T_LAST_POST DESC"

			set rs = my_Conn.Execute (strSql)
			
			if not rs.eof then
				strLast_Post = rs("T_LAST_POST")
				strLast_Post_Author = rs("T_LAST_POST_AUTHOR")
			else
				strLast_Post = ""
				strLast_Post_Author = ""
			end if
			
			rs.Close
			set rs = nothing

			'## Forum_SQL - Decrease count of total replies in Forum by 1
			strSql =  "UPDATE " & strTablePrefix & "FORUM "
			strSql = strSql & " SET " & strTablePrefix & "FORUM.F_COUNT = " & strTablePrefix & "FORUM.F_COUNT - " & 1 & " "
			if strLast_Post <> "" then 
				strSql = strSql & ", F_LAST_POST = '" & strLast_Post & "'"
				if strLast_Post_Author <> "" then 



					strSql = strSql & ", F_LAST_POST_AUTHOR = " & strLast_Post_Author



				end if
			end if
			strSql = strSql & " WHERE " & strTablePrefix & "FORUM.FORUM_ID = " & Request.Form("FORUM_ID")

			my_Conn.Execute strSql

			'## FORUM_SQL - Decrease count of total replies in Totals table by 1
			strSql = "UPDATE " & strTablePrefix & "TOTALS "
			strSql = strSql & " SET " & strTablePrefix & "TOTALS.P_COUNT = " & strTablePrefix & "TOTALS.P_COUNT - " & 1 & " "

			my_Conn.Execute strSql
%>
<P><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">回复已删除！</font></p>
<P>(<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">记得重新整理页面。</font>)</p>
<%		Else %>
<P><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">你不是管理员不能删除回复</p>
<p><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript: onClick= history.go(-1)">返回重新认证</a></font></p>
<%		end if %>
<%	Else %>
<P><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">你不是管理员不能删除回复</p>
<p><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript: onClick= history.go(-1)">返回重新认证</a></font></p>
<%	end if

else
	if Request.QueryString("mode") = "DeleteTopic" then
		mLev = cint(ChkUser2(STRdbntUserName, Request.Form("Pass"))) 
		if mLev > 0 then  '## is Member
			if (chkForumModerator(Request.Form("FORUM_ID"), STRdbntUserName) = "1") or (mLev = 4) then
				delAr = split(Request.Form("TOPIC_ID"), ",")
				for i = 0 to ubound(delAr) 

					'## Forum_SQL - count total number of replies of TOPIC_ID  in Reply table



					set rs = Server.CreateObject("ADODB.Recordset")
					strSql = "SELECT count(" & strTablePrefix & "REPLY.REPLY_ID) AS cnt "
					strSql = strSql & " FROM " & strTablePrefix & "REPLY "
					strSql = strSql & " WHERE " & strTablePrefix & "REPLY.TOPIC_ID = " & cint(delAr(i))
					rs.Open strSql, my_Conn
					risposte = rs("cnt")
					rs.close
					set rs = nothing

					'## Forum_SQL - Delete the actual topics
					strSql = "DELETE FROM " & strTablePrefix & "TOPICS "
					strSql = strSql & " WHERE " & strTablePrefix & "TOPICS.TOPIC_ID = " & cint(delAr(i))
					my_Conn.Execute strSql

					'## Forum_SQL - Delete all replys related to the topics
					strSql = "DELETE FROM " & strTablePrefix & "REPLY "
					strSql = strSql & " WHERE " & strTablePrefix & "REPLY.TOPIC_ID = " & cint(delAr(i))

					my_Conn.Execute strSql	'## Forum_SQL - Get last_post and last_post_author for Forum
					strSql = "SELECT T_LAST_POST, T_LAST_POST_AUTHOR"
					strSql = strSql & " FROM " & strTablePrefix & "TOPICS "			
					strSql = strSql & " WHERE FORUM_ID = " & Request.Form("FORUM_ID") & " "
					strSql = strSql & " ORDER BY T_LAST_POST DESC"
					
					set rs = my_Conn.Execute (strSql)
			
					if not rs.eof then
						rs.movefirst
						strLast_Post = rs("T_LAST_POST")
						strLast_Post_Author = rs("T_LAST_POST_AUTHOR")
					else
						strLast_Post = ""
						strLast_Post_Author = ""
					end if
			
					rs.Close
					set rs = nothing

					'## Forum_SQL - Update count of replies to a topic in Forum table
					strSql = "UPDATE " & strTablePrefix & "FORUM "
					strSql = strSql & " SET " & strTablePrefix & "FORUM.F_COUNT = " & strTablePrefix & "FORUM.F_COUNT - " & cint(risposte) + 1
					strSql = strSql & " ,   " & strTablePrefix & "FORUM.F_TOPICS = " & strTablePrefix & "FORUM.F_TOPICS - "	& 1				
					if strLast_Post <> "" then 						
						strSql = strSql & ", F_LAST_POST = '" & strLast_Post & "' "
						if strLast_Post_Author <> "" then
							strSql = strSql & ", F_LAST_POST_AUTHOR = " & strLast_Post_Author
						end if
					else
					strSql = strSql & ", F_LAST_POST = F_CREATED"
					strSql = strSql & ", F_LAST_POST_AUTHOR = NULL" 
					end if

					strSql = strSql & " WHERE " & strTablePrefix & "FORUM.FORUM_ID = " & Request.Form("FORUM_ID")
					my_Conn.Execute strSql  
					'## Forum_SQL - Update total TOPICS in Totals table

					strSql = "UPDATE " & strTablePrefix & "TOTALS "
					strSql = strSql & " SET " & strTablePrefix & "TOTALS.T_COUNT = " & strTablePrefix & "TOTALS.T_COUNT - " & 1
					strSql = strSql & ",    " & strTablePrefix & "TOTALS.P_COUNT = " & strTablePrefix & "TOTALS.P_COUNT - " & cint(risposte) + 1
					my_Conn.Execute strSql					

				next
%>
<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><b>主题已删除！</b></font></p>
<%			Else %>
<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><b>你不是管理员不能删除主题</font><br>
<br>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript: onClick= history.go(-1) ">返回重新认证</a></font></p>
<%			end if %>	  
<%		Else %>
<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><b>你不是管理员不能删除主题</font><br>
<br>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript: onClick= history.go(-1)">返回重新认证</a></font></p>
<%
		end if 
	else 
		if Request.QueryString("mode") = "DeleteForum" then
			mLev = cint(ChkUser2(strDBNTUserName, Request.Form("Pass"))) 
			if mLev > 0 then  '## is Member
				if mLev = 4 then
					delAr = split(Request.Form("FORUM_ID"), ",")
					for i = 0 to ubound(delAr) 
						'## Forum_SQL - Delete all replys related to the topics
						strSql = "DELETE FROM " & strTablePrefix & "REPLY "
						strSql = strSql & " WHERE FORUM_ID = " & cint(delAr(i))

						my_Conn.Execute strSql

						'## Forum_SQL - Delete the actual topics
						strSql = "DELETE FROM " & strTablePrefix & "TOPICS "
						strSql = strSql & " WHERE " & strTablePrefix & "TOPICS.FORUM_ID = " & cint(delAr(i))

						my_Conn.Execute strSql

						'## Forum_SQL - Delete the moderators
						strSql = "DELETE FROM " & strTablePrefix & "MODERATOR "
						strSql = strSql & " WHERE " & strTablePrefix & "MODERATOR.FORUM_ID = " & cint(delAr(i))

						my_Conn.Execute strSql

						'## Forum_SQL - Delete the actual forums
						strSql = "DELETE FROM " & strTablePrefix & "FORUM "
						strSql = strSql & " WHERE " & strTablePrefix & "FORUM.FORUM_ID = " & cint(delAr(i))

						my_Conn.Execute strSql

						
						'## Forum_SQL - count total number of replies in Reply table

						set rs = Server.CreateObject("ADODB.Recordset")

						strSql = "SELECT count(" & strTablePrefix & "REPLY.REPLY_ID) AS cnt "
						strSql = strSql & " FROM " & strTablePrefix & "REPLY "
						
						rs.Open strSql, my_Conn
						risreply = rs("cnt")
						rs.close

						set rs = nothing

						set rs = Server.CreateObject("ADODB.Recordset")

						'## Forum_SQL - count total number of Topics in Topics table
						strSql = "SELECT count(" & strTablePrefix & "TOPICS.TOPIC_ID) AS cnt "
						strSql = strSql & " FROM " & strTablePrefix & "TOPICS "

						rs.Open strSql, my_Conn
						rispost = rs("cnt")
						rs.close

						set rs = nothing

						'## Forum_SQL - Update total topics and posts in Totals table
						strSql = "UPDATE " & strTablePrefix & "TOTALS "
						strSql = strSql & " SET " & strTablePrefix & "TOTALS.P_COUNT = " & cint(risreply + rispost)
						strSql = strSql & ",    " & strTablePrefix & "TOTALS.T_COUNT = " & cint(rispost)

						my_Conn.Execute strSql
					next
%>
<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><b>论坛已删除！</b></font></p>
<%				Else %>
<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><b>你不是管理员不能删除论坛</font><br>
<br>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript: onClick= history.go(-1) ">返回重新认证</a></font></p>
<%				end if %>	  
<%			Else %>
<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><b>你不是管理员不能删除论坛</font><br>
<br>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript: onClick= history.go(-1)">返回重新认证</a></font></p>
<%
			end if 
		else
			if Request.QueryString("mode") = "DeleteCategory" then
				mLev = cint(ChkUser2(strDBNTUserName, Request.Form("Pass"))) 
				if mLev > 0 then  '## is Member
					if mLev = 4 then
						delAr = split(Request.Form("CAT_ID"), ",")
						for i = 0 to ubound(delAr) 
							'## Forum_SQL - Delete all replys related to the topics
							strSql = "DELETE FROM " & strTablePrefix & "REPLY "
							strSql = strSql & " WHERE " & strTablePrefix & "REPLY.CAT_ID = " & cint(delAr(i))

							my_Conn.Execute strSql
							
							'## Forum_SQL - Delete the actual topics
							strSql = "DELETE FROM " & strTablePrefix & "TOPICS "
							strSql = strSql & " WHERE " & strTablePrefix & "TOPICS.CAT_ID = " & cint(delAr(i))

							my_Conn.Execute strSql

							'## Forum_SQL - Delete the actual forums
							strSql = "DELETE FROM " & strTablePrefix & "FORUM "
							strSql = strSql & " WHERE " & strTablePrefix & "FORUM.CAT_ID = " & cint(delAr(i))

							my_Conn.Execute strSql

							'## Forum_SQL - Delete the actual category
							strSql = "DELETE FROM " & strTablePrefix & "CATEGORY "
							strSql = strSql & " WHERE " & strTablePrefix & "CATEGORY.CAT_ID = " & cint(delAr(i))

							my_Conn.Execute strSql
							
							
							'## Forum_SQL - count total number of replies in Reply table
							set rs = Server.CreateObject("ADODB.Recordset")

							strSql = "SELECT count(" & strTablePrefix & "REPLY.REPLY_ID) AS cnt "
							strSql = strSql & " FROM " & strTablePrefix & "REPLY "

							rs.Open strSql, my_Conn
							risreply = rs("cnt")
							rs.close
	
							set rs = nothing

						
							'## Forum_SQL - count total number of Topics in Topics table
							set rs = Server.CreateObject("ADODB.Recordset")

							strSql = "SELECT count(" & strTablePrefix & "TOPICS.TOPIC_ID) AS cnt "
							strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
							
							rs.Open strSql, my_Conn
							rispost = rs("cnt")
							rs.close
							
							P_Count = cint(risreply + rispost)
							T_Count = cint(rispost)
							'## Forum_SQL - Update total topics and posts in Totals table
							strSql = "UPDATE " & strTablePrefix & "TOTALS "
							strSql = strSql & " SET " & strTablePrefix & "TOTALS.P_COUNT = " & P_Count & " "
							strSql = strSql & ",    " & strTablePrefix & "TOTALS.T_COUNT = " & T_Count & " "
							my_Conn.Execute strSql
						next
%>
<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><b>分类已删除！</b></font></p>
<%					else %>
<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><b>你不是管理员不能删除分类</font><br>
<br>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript: onClick= history.go(-1) ">返回重新认证</a></font></p>
<%					end if %>	  
<%				else %>
<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><b>你不是管理员不能删除分类</font><br>
<br>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript: onClick= history.go(-1)">返回重新认证</a></font></p>
<%
				end if 
			else
				if Request.QueryString("mode") = "DeleteMember" then
					mLev = cint(ChkUser2(STRdbntUserName, Request.Form("Pass"))) 
					if mLev > 0 then  '## is Member
						if mLev = 4 then

							'## Forum_SQL - Remove the member from the moderator table
							strSql = "DELETE FROM " & strTablePrefix & "MODERATOR "
							strSql = strSql & " WHERE " & strTablePrefix & "MODERATOR.MEMBER_ID = " & Request.Form("MEMBER_ID")

							my_Conn.Execute (strSql)

							'## Forum_SQL - Select postcount
							strSql = "SELECT COUNT(T_AUTHOR) AS POSTCOUNT "
							strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
							strSql = strSql & " WHERE T_AUTHOR = " & Request.Form("MEMBER_ID")

							set rs = my_Conn.Execute (strSql)
							if not rs.eof then
								intPostcount = rs("POSTCOUNT")
							else
								intPostcount = 0
							end if
							
							rs.close

							
							'## Forum_SQL - Select postcount
							strSql = "SELECT COUNT(R_AUTHOR) AS REPLYCOUNT "
							strSql = strSql & " FROM " & strTablePrefix & "REPLY "
							strSql = strSql & " WHERE R_AUTHOR = " & Request.Form("MEMBER_ID")

							set rs = my_Conn.Execute (strSql)
							if not rs.eof then
								intReplycount = rs("REPLYCOUNT")
							else
								intReplycount = 0
							end if
				
							rs.close
							
							if ((intReplycount + intPostCount) = 0) then

								'## Forum_SQL - Delete the Member - Member has no posts
								strSql = "DELETE FROM " & strMemberTablePrefix & "MEMBERS "
								strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & Request.Form("MEMBER_ID")

								my_Conn.Execute strSql

							else
																				
								'## Forum_SQL - disable account - Member has posts, cannot delete just disable account
								strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
								strSql = strSql & " SET " & strMemberTablePrefix & "MEMBERS.M_STATUS = " & 0
								strSql = strSql & ",    " & strMemberTablePrefix & "MEMBERS.M_EMAIL = ' '"
								strSql = strSql & ",    " & strMemberTablePrefix & "MEMBERS.M_LEVEL = " & 1
								strSql = strSql & ",    " & strMemberTablePrefix & "MEMBERS.M_NAME = 'n/a'"
								strSql = strSql & ",    " & strMemberTablePrefix & "MEMBERS.M_COUNTRY = ' '"
								strSql = strSql & ",    " & strMemberTablePrefix & "MEMBERS.M_TITLE = 'deleted'"
								strSql = strSql & ",    " & strMemberTablePrefix & "MEMBERS.M_HOMEPAGE = ' '"
								strSql = strSql & ",    " & strMemberTablePrefix & "MEMBERS.M_AIM = ' '"
								strSql = strSql & ",    " & strMemberTablePrefix & "MEMBERS.M_YAHOO = ' '"
								strSql = strSql & ",    " & strMemberTablePrefix & "MEMBERS.M_ICQ = ' '"
								strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.MEMBER_ID = " & Request.Form("MEMBER_ID")

								my_Conn.Execute strSql

							end if

							

							'## Forum_SQL - Update total of Members in Totals table
							strSql = "UPDATE " & strTablePrefix & "TOTALS "
							strSql = strSql & " SET " & strTablePrefix & "TOTALS.U_COUNT = " & strTablePrefix & "TOTALS.U_COUNT - " & 1

							my_Conn.Execute strSql
%>
<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><b>会员已删除！</b></font></p>
<%						else %>
<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><b>你不是管理员不能删除会员</font><br>
<br>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript: onClick= history.go(-1) ">返回重新认证</a></font></p>
<%						end if %>	  
<%					else %>
<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>"><b>你不是管理员不能删除会员</font><br>
<br>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript: onClick= history.go(-1)">返回重新认证</a></font></p>
<%
					end if 
				else
%>

<P><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">Delete 
<%				if Request.QueryString("mode") = "Member" then %>
Member
<%				else %>
<%					if Request.QueryString("mode") = "Category" then %>
Category
<%					else %>
<%						if Request.QueryString("mode") = "Forum" then %>
Forum
<%						else %>
<%							if Request.QueryString("mode") = "Topic" then %>
Topic
<%							else %>
<%								if Request.QueryString("mode") = "Reply" then %>
Reply
<%								end if %>
<%							end if %>
<%						end if %>
<%					end if %>
<%				end if %>
</font></p>

<p><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b><font color=red>注意：</font></b>  
<%				if Request.QueryString("mode") = "Member" then %>
只有管理员才能删除会员。
<%				else %>
<%					if Request.QueryString("mode") = "Category" then %>
只有管理员才能删除分类。
<%					else %>
<%						if Request.QueryString("mode") = "Forum" then %>
只有管理员才能删除论坛。
<%						else %>
<%							if Request.QueryString("mode") = "Topic" then %>
只有版主和管理员才能删除主题。
<%							else %>
<%								if Request.QueryString("mode") = "Reply" then %>
只有版主和管理者才能删除回复。
<%								end if %>
<%							end if %>
<%						end if %>
<%					end if %>
<%				end if %>
</font></p>

<form action="pop_delete.asp?mode=<% if Request.QueryString("mode") = "Member" then Response.Write("DeleteMember")%><% if Request.QueryString("mode") = "Category" then Response.Write("DeleteCategory")%><% if Request.QueryString("mode") = "Forum" then Response.Write("DeleteForum")%><% if Request.QueryString("mode") = "Topic" then Response.Write("DeleteTopic")%><%if Request.QueryString("mode") = "Reply" then Response.Write("DeleteReply")%>" method=post id=Form1 name=Form1>
<input type=hidden name="REPLY_ID" value="<% =Request.QueryString("REPLY_ID") %>">
<input type=hidden name="TOPIC_ID" value="<% =Request.QueryString("TOPIC_ID") %>">
<input type=hidden name="FORUM_ID" value="<% =Request.QueryString("FORUM_ID") %>">
<input type=hidden name="CAT_ID" value="<% =Request.QueryString("CAT_ID") %>">
<input type=hidden name="MEMBER_ID" value="<% =Request.QueryString("MEMBER_ID") %>">

<table border="0" width="75%" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="<% =strPopUpBorderColor %>">
    <table border="0" width="100%" cellspacing="1" cellpadding="1">
<%				if strAuthType="db" then %>
      <tr>
        <td bgColor=<% =strPopUpTableColor %> align=right nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">用户名：</font></b></td>
        <td bgColor=<% =strPopUpTableColor %>><input type=text name="User" value="<% =Request.Cookies(strUniqueID & "User")("Name")%>" size=20></td>
      </tr>
      <tr>
        <td bgColor=<% =strPopUpTableColor %> align=right nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">密码：</FONT></b></td>
        <td bgColor=<% =strPopUpTableColor %>><input type=Password name="Pass" value="<% =Request.Cookies(strUniqueID & "User")("Pword")%>" size=20></td>
      </tr>
<%				else %>
<%					if strAuthType="nt" then %>
      <tr>
        <td bgColor=<% =strPopUpTableColor %> align=right nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">NT 帐号：</font></b></td>
        <td bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><%=Session(strCookieURL & "userid")%></font></td>
      </tr>
<%					end if %>
<%				end if %>
      <tr>
        <td bgColor=<% =strPopUpTableColor %> colspan=2 align=center><Input type=Submit value="提交" id=Submit1 name=Submit1></td>
      </tr>
    </table>
    </td>
  </tr>
</table>
</form>
<%
				end if
			end if
		end if 
	end if 
end if 
%>
<!--#INCLUDE FILE="inc_footer_short.asp" -->