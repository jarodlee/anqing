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

on error resume next
	
set my_Conn = Server.CreateObject("ADODB.Connection")
my_Conn.Open strConnString	
	
for counter = 0 to my_conn.Errors.Count -1
	ConnErrorNumber = my_conn.Errors(counter).Number
	if ConnErrorNumber <> 0 then 
		
		Response.write ConnErrorNumber & " : " & my_conn.Errors(counter).Description & "<br>" 
		
	end if
	my_conn.Errors.Clear 
next
	
my_Conn.Errors.Clear
on error resume next

strSql = "DROP TABLE IF EXISTS " & strTablePrefix & "CATEGORY "

my_Conn.Execute strSql
ChkDBInstall()

strSql = "CREATE TABLE " & strTablePrefix & "CATEGORY ( "
strSql = strSql & "CAT_ID INT (11) DEFAULT '' NOT NULL auto_increment, "
strSql = strSql & "CAT_STATUS SMALLINT (6) DEFAULT '1' NOT NULL , "
strSql = strSql & "CAT_NAME VARCHAR (100),  "
strSql = strSql & "PRIMARY KEY (CAT_ID), "
strSql = strSql & "KEY " & strTablePrefix & "CATEGORY_CAT_ID(CAT_ID), "
strSql = strSql & "KEY " & strTablePrefix & "CATEGORY_CAT_STATUS (CAT_STATUS) )"

my_Conn.Execute strSql
ChkDBInstall()

strSql = "DROP TABLE IF EXISTS " & strTablePrefix & "CONFIG "

my_Conn.Execute strSql
ChkDBInstall()

strSql = "CREATE TABLE " & strTablePrefix & "CONFIG ("
strSql = strSql & "CONFIG_ID INT (11) DEFAULT '' NOT NULL auto_increment , "
strSql = strSql & "C_STRVERSION	VARCHAR (50) DEFAULT '" & strNewVersion & "' , "
strSql = strSql & "C_STRFORUMTITLE VARCHAR (255) DEFAULT 'Snitz Forums 2000' , "
strSql = strSql & "C_STRCOPYRIGHT VARCHAR (50) DEFAULT '&copy; 2000 Snitz Communications' , "
strSql = strSql & "C_STRTITLEIMAGE VARCHAR (255) DEFAULT  'logo_snitz_forums_2000.gif' , "
strSql = strSql & "C_STRHOMEURL VARCHAR (255) DEFAULT  '../' , "
strSql = strSql & "C_STRFORUMURL VARCHAR (255) DEFAULT  'http://snitz.forum.com/forum/' , "
strSql = strSql & "C_STRAUTHTYPE VARCHAR (50) DEFAULT  'db' , "
strSql = strSql & "C_STREMAIL smallint (6) DEFAULT  '1' , "
strSql = strSql & "C_STRUNIQUEEMAIL smallint (6) DEFAULT  '1' , "
strSql = strSql & "C_STRMAILMODE VARCHAR (100) DEFAULT  'cdonts' , "
strSql = strSql & "C_STRMAILSERVER VARCHAR (255) DEFAULT  'your.mailserver.com' , "
strSql = strSql & "C_STRSENDER VARCHAR (255) DEFAULT  'your.email@yourserver.com' , "
strSql = strSql & "C_STRSETCOOKIETOFORUM smallint (6) DEFAULT '1' , "
strSql = strSql & "C_STRDATETYPE VARCHAR (50) DEFAULT  'mdy' , "
strSql = strSql & "C_STRTIMETYPE VARCHAR (50) DEFAULT  '24' , "
strSql = strSql & "C_STRTIMEADJUSTLOCATION VARCHAR (50) DEFAULT  '0' , "
strSql = strSql & "C_STRTIMEADJUST int (11) DEFAULT  '0' , "
strSql = strSql & "C_STRMOVETOPICMODE smallint (6) DEFAULT  '1' , "
strSql = strSql & "C_STRPRIVATEFORUMS smallint (6) DEFAULT  '0' , "
strSql = strSql & "C_STRSHOWMODERATORS smallint (6) DEFAULT '1' , "
strSql = strSql & "C_STRSHOWRANK smallint (6) DEFAULT '1' , "
strSql = strSql & "C_STRHIDEEMAIL smallint (6) DEFAULT '0' , "
strSql = strSql & "C_STRIPLOGGING smallint (6) DEFAULT '1' , "
strSql = strSql & "C_STRALLOWFORUMCODE smallint (6) DEFAULT  '1', "
strSql = strSql & "C_STRIMGINPOSTS smallint (6) DEFAULT '0' , "
strSql = strSql & "C_STRALLOWHTML smallint (6) DEFAULT '0' , "
strSql = strSql & "C_STRSECUREADMIN smallint (6) DEFAULT '1' , "
strSql = strSql & "C_STRNOCOOKIES smallint (6) DEFAULT '0' , "
strSql = strSql & "C_STREDITEDBYDATE smallint (6) DEFAULT '1' , "
strSql = strSql & "C_STRHOTTOPIC smallint (6) DEFAULT '1' , "
strSql = strSql & "C_INTHOTTOPICNUM int (11) DEFAULT '20' , "
strSql = strSql & "C_STRHOMEPAGE smallint (6) DEFAULT '1' , "
strSql = strSql & "C_STRAIM smallint (6) DEFAULT '1' , "
strSql = strSql & "C_STRYAHOO smallint (6) DEFAULT '1' , "
strSql = strSql & "C_STRICQ smallint (6) DEFAULT '1' , "
strSql = strSql & "C_STRICONS smallint (6) DEFAULT '1' , "
strSql = strSql & "C_STRGFXBUTTONS smallint (6) DEFAULT '1' , "
strSql = strSql & "C_STRBADWORDFILTER smallint (6) DEFAULT '1' , "
strSql = strSql & "C_STRBADWORDS VARCHAR (255) DEFAULT 'fuck|wank|shit|pussy|cunt' , "
strSql = strSql & "C_STRDEFAULTFONTFACE VARCHAR (255) DEFAULT '宋体, Arial, Helvetica' , "
strSql = strSql & "C_STRDEFAULTFONTSIZE VARCHAR (50) DEFAULT '2' , "
strSql = strSql & "C_STRHEADERFONTSIZE VARCHAR (50) DEFAULT '4' , "
strSql = strSql & "C_STRFOOTERFONTSIZE VARCHAR (50) DEFAULT '1' , "
strSql = strSql & "C_STRPAGEBGCOLOR VARCHAR (50) DEFAULT 'white' , "
strSql = strSql & "C_STRDEFAULTFONTCOLOR VARCHAR (50) DEFAULT 'midnightblue' , "
strSql = strSql & "C_STRLINKCOLOR VARCHAR (50) DEFAULT 'darkblue' , "
strSql = strSql & "C_STRLINKTEXTDECORATION VARCHAR (50) DEFAULT 'underline' , "
strSql = strSql & "C_STRVISITEDLINKCOLOR VARCHAR (50) DEFAULT 'blue', "
strSql = strSql & "C_STRVISITEDTEXTDECORATION VARCHAR (50) DEFAULT 'underline'  , "
strSql = strSql & "C_STRACTIVELINKCOLOR VARCHAR (50) DEFAULT 'red' , "
strSql = strSql & "C_STRHOVERFONTCOLOR VARCHAR (50) DEFAULT 'red' , "
strSql = strSql & "C_STRHOVERTEXTDECORATION VARCHAR (50) DEFAULT 'underline' , "
strSql = strSql & "C_STRHEADCELLCOLOR VARCHAR (50) DEFAULT 'midnightblue' , "
strSql = strSql & "C_STRHEADFONTCOLOR VARCHAR (50) DEFAULT 'mintcream' , "
strSql = strSql & "C_STRCATEGORYCELLCOLOR VARCHAR (50) DEFAULT 'slateblue' , "
strSql = strSql & "C_STRCATEGORYFONTCOLOR VARCHAR (50) DEFAULT 'mintcream' , "
strSql = strSql & "C_STRFORUMFIRSTCELLCOLOR VARCHAR (50) DEFAULT 'whitesmoke' , "
strSql = strSql & "C_STRFORUMCELLCOLOR VARCHAR (50) DEFAULT 'whitesmoke' , "
strSql = strSql & "C_STRALTFORUMCELLCOLOR VARCHAR (50) DEFAULT 'gainsboro' , "
strSql = strSql & "C_STRFORUMFONTCOLOR VARCHAR (50) DEFAULT 'midnightblue' , "
strSql = strSql & "C_STRFORUMLINKCOLOR VARCHAR (50) DEFAULT 'darkblue' , "
strSql = strSql & "C_STRTABLEBORDERCOLOR VARCHAR (50) DEFAULT 'black' , "
strSql = strSql & "C_STRPOPUPTABLECOLOR VARCHAR (50) DEFAULT 'navyblue' , "
strSql = strSql & "C_STRPOPUPBORDERCOLOR VARCHAR (50) DEFAULT 'black' , "
strSql = strSql & "C_STRNEWFONTCOLOR VARCHAR (50) DEFAULT 'blue' , "
strSql = strSql & "C_STRTOPICWIDTHLEFT VARCHAR (10) DEFAULT '100' , "
strSql = strSql & "C_STRTOPICNOWRAPLEFT smallint (6) DEFAULT '1' , "
strSql = strSql & "C_STRTOPICWIDTHRIGHT VARCHAR (10) DEFAULT '100%' , "
strSql = strSql & "C_STRTOPICNOWRAPRIGHT smallint (6)  DEFAULT '0' , "
strSql = strSql & "C_STRRANKADMIN VARCHAR (50) DEFAULT 'Administrator' , "
strSql = strSql & "C_STRRANKMOD VARCHAR (50) DEFAULT 'Moderator' , "
strSql = strSql & "C_STRRANKLEVEL0 VARCHAR (50) DEFAULT 'Starting Member' , "
strSql = strSql & "C_STRRANKLEVEL1 VARCHAR (50) DEFAULT 'New Member',  "
strSql = strSql & "C_STRRANKLEVEL2 VARCHAR (50) DEFAULT 'Junior Member' , "
strSql = strSql & "C_STRRANKLEVEL3 VARCHAR (50) DEFAULT 'Average Member' , "
strSql = strSql & "C_STRRANKLEVEL4 VARCHAR (50) DEFAULT 'Senior Member' , "
strSql = strSql & "C_STRRANKLEVEL5 VARCHAR (50) DEFAULT 'Advanced Member' , "
strSql = strSql & "C_STRRANKCOLORADMIN VARCHAR (50)DEFAULT 'gold' , "
strSql = strSql & "C_STRRANKCOLORMOD VARCHAR (50) DEFAULT 'silver' , "
strSql = strSql & "C_STRRANKCOLOR0 VARCHAR (50) DEFAULT 'bronze' , "
strSql = strSql & "C_STRRANKCOLOR1 VARCHAR (50) DEFAULT 'bronze' , "
strSql = strSql & "C_STRRANKCOLOR2 VARCHAR (50) DEFAULT 'bronze' , "
strSql = strSql & "C_STRRANKCOLOR3 VARCHAR (50) DEFAULT 'bronze' , "
strSql = strSql & "C_STRRANKCOLOR4 VARCHAR (50) DEFAULT 'bronze' , "
strSql = strSql & "C_STRRANKCOLOR5 VARCHAR (50) DEFAULT 'bronze' , "
strSql = strSql & "C_INTRANKLEVEL0 smallint (6) DEFAULT '0' , "
strSql = strSql & "C_INTRANKLEVEL1 smallint (6) DEFAULT '50' , "
strSql = strSql & "C_INTRANKLEVEL2 smallint (6) DEFAULT '100' , "
strSql = strSql & "C_INTRANKLEVEL3 smallint (6) DEFAULT '500' , "
strSql = strSql & "C_INTRANKLEVEL4 smallint (6) DEFAULT '1000' , "
strSql = strSql & "C_INTRANKLEVEL5 smallint (6) DEFAULT '2000' , "
strSql = strSql & "C_STRSIGNATURES smallint (6) DEFAULT '1',   "
strSql = strSql & "C_STRSHOWSTATISTICS smallint (6) DEFAULT '1', "
strSql = strSql & "C_STRSHOWIMAGEPOWEREDBY smallint (6) DEFAULT '1', "
strSql = strSql & "C_STRLOGONFORMAIL smallint (6) DEFAULT '1', "
strSql = strSql & "C_STRSHOWPAGING smallint (6) DEFAULT '0', "
strSql = strSql & "C_STRSHOWTOPICNAV smallint (6) DEFAULT '0', "
strSql = strSql & "C_STRPAGESIZE smallint (6) DEFAULT '15', "
strSql = strSql & "C_STRPAGENUMBERSIZE smallint (6) DEFAULT '10', "
strSql = strSql & "C_STRFULLNAME smallint (6) DEFAULT '0', "
strSql = strSql & "C_STRPICTURE smallint (6) DEFAULT '0', "
strSql = strSql & "C_STRSEX smallint (6) DEFAULT '0', "
strSql = strSql & "C_STRCITY smallint (6) DEFAULT '0', "
strSql = strSql & "C_STRSTATE smallint (6) DEFAULT '0', "
strSql = strSql & "C_STRAGE smallint (6) DEFAULT '0', "
strSql = strSql & "C_STRCOUNTRY smallint (6) DEFAULT '0', "
strSql = strSql & "C_STROCCUPATION smallint (6) DEFAULT '0', "
strSql = strSql & "C_STRBIO smallint (6) DEFAULT '0', "
strSql = strSql & "C_STRHOBBIES smallint (6) DEFAULT '0', "
strSql = strSql & "C_STRLNEWS smallint (6) DEFAULT '0', "
strSql = strSql & "C_STRQUOTE smallint (6) DEFAULT '0', "
strSql = strSql & "C_STRMARSTATUS smallint (6) DEFAULT '0', "
strSql = strSql & "C_STRFAVLINKS smallint (6) DEFAULT '1', "
strSql = strSql & "C_STRRECENTTOPICS smallint (6) DEFAULT '0', "
strSql = strSql & "C_STRAUTOLOGON smallint (6) DEFAULT '0', "
strSql = strSql & "C_STRNTGROUPS smallint (6) DEFAULT '0', "
strSql = strSql & "PRIMARY KEY (CONFIG_ID)) "

my_Conn.Execute strSql
ChkDBInstall()

strSql = "DROP TABLE IF EXISTS " & strTablePrefix & "CONFIG_NEW "

my_Conn.Execute strSql
ChkDBInstall()

strSql = "CREATE TABLE " & strTablePrefix & "CONFIG_NEW ( "
strSql = strSql & "ID int (11) DEFAULT '' NOT NULL auto_increment, "
strSql = strSql & "C_VARIABLE VARCHAR (255) , "
strSql = strSql & "C_VALUE VARCHAR (255),  "
strSql = strSql & "PRIMARY KEY (ID)) "

my_Conn.Execute strSql
ChkDBInstall()

strSql = "DROP TABLE IF EXISTS " & strTablePrefix & "FORUM "

my_Conn.Execute strSql
ChkDBInstall()

strSql = "CREATE TABLE " & strTablePrefix & "FORUM ( "
strSql = strSql & "CAT_ID int (11) DEFAULT '1' NOT NULL , "
strSql = strSql & "FORUM_ID smallint (6) DEFAULT '0' NOT NULL auto_increment, "
strSql = strSql & "F_STATUS smallint (6) DEFAULT '1', "
strSql = strSql & "F_MAIL smallint (6) DEFAULT '1' , "
strSql = strSql & "F_SUBJECT VARCHAR (100) DEFAULT ' ' , "
strSql = strSql & "F_URL VARCHAR (255) DEFAULT ' ' , "
strSql = strSql & "F_DESCRIPTION VARCHAR (255) DEFAULT ' ' , "
strSql = strSql & "F_TOPICS int (11) DEFAULT '0' , "
strSql = strSql & "F_COUNT int (11) DEFAULT '0' , "
strSql = strSql & "F_LAST_POST VARCHAR (50) DEFAULT ' ' , "
strSql = strSql & "F_PASSWORD_NEW VARCHAR (255) DEFAULT ' ' , "
strSql = strSql & "F_USERLIST VARCHAR (255) DEFAULT ' ' , "
strSql = strSql & "F_PRIVATEFORUMS int (11) DEFAULT '0' , "
strSql = strSql & "F_TYPE smallint (6) DEFAULT '0' , "
strSql = strSql & "F_IP VARCHAR (50) DEFAULT '000.000.000.000' ,  "
strSql = strSql & "F_LAST_POST_AUTHOR int (11) DEFAULT '1' ,  "
strSql = strSql & "PRIMARY KEY (CAT_ID, FORUM_ID), "
strSql = strSql & "KEY " & strTablePrefix & "FORUM_FORUM_ID(FORUM_ID), "
strSql = strSql & "KEY " & strTablePrefix & "FORUM_CAT_ID(CAT_ID)) "


my_Conn.Execute strSql
ChkDBInstall()

strSql = "DROP TABLE IF EXISTS " & strMemberTablePrefix & "MEMBERS "

my_Conn.Execute strSql
ChkDBInstall()

strSql = "CREATE TABLE " & strMemberTablePrefix & "MEMBERS ( "
strSql = strSql & "MEMBER_ID int (11) DEFAULT '' NOT NULL AUTO_INCREMENT, "
strSql = strSql & "M_STATUS smallint (6) DEFAULT '0' , "
strSql = strSql & "M_NAME VARCHAR (75) DEFAULT '' , "
strSql = strSql & "M_USERNAME VARCHAR (150) DEFAULT '' , "
strSql = strSql & "M_PASSWORD VARCHAR (25) DEFAULT '' , "
strSql = strSql & "M_EMAIL VARCHAR (50) DEFAULT '' , "
strSql = strSql & "M_COUNTRY VARCHAR (20) DEFAULT '' , "
strSql = strSql & "M_HOMEPAGE VARCHAR (50) DEFAULT '' , "
strSql = strSql & "M_SIG VARCHAR (255) DEFAULT '' , "
strSql = strSql & "M_DEFAULT_VIEW int (11)  DEFAULT '1' , "
strSql = strSql & "M_LEVEL smallint (6)  DEFAULT '1' , "
strSql = strSql & "M_AIM VARCHAR (150)  DEFAULT '' , "
strSql = strSql & "M_YAHOO VARCHAR (150)  DEFAULT '' , "
strSql = strSql & "M_ICQ VARCHAR (150)  DEFAULT '' , "
strSql = strSql & "M_POSTS int (11)  DEFAULT '0' , "
strSql = strSql & "M_DATE VARCHAR (50)  DEFAULT '' , "
strSql = strSql & "M_LASTHEREDATE VARCHAR (50)  DEFAULT '' , "
strSql = strSql & "M_LASTPOSTDATE VARCHAR (50)  DEFAULT '' , "
strSql = strSql & "M_TITLE VARCHAR (50)  DEFAULT '' , "
strSql = strSql & "M_SUBSCRIPTION smallint (6)  DEFAULT '0' , "
strSql = strSql & "M_HIDE_EMAIL smallint (6) DEFAULT '0' , "
strSql = strSql & "M_RECEIVE_EMAIL smallint (6)  DEFAULT '0' , "
strSql = strSql & "M_LAST_IP VARCHAR (50)  DEFAULT '000.000.000.000' , "
strSql = strSql & "M_IP VARCHAR (50)  DEFAULT '000.000.000.000' , "
strSql = strSql & "M_FIRSTNAME VARCHAR (100) DEFAULT '' , "
strSql = strSql & "M_LASTNAME VARCHAR (100) DEFAULT '' , "
strSql = strSql & "M_OCCUPATION VARCHAR (255) DEFAULT '' , "
strSql = strSql & "M_SEX VARCHAR (50) DEFAULT '' , "
strSql = strSql & "M_AGE VARCHAR (10) DEFAULT '' , "
strSql = strSql & "M_HOBBIES TEXT , "
strSql = strSql & "M_LNEWS TEXT , "
strSql = strSql & "M_QUOTE TEXT , "
strSql = strSql & "M_BIO TEXT , "
strSql = strSql & "M_MARSTATUS VARCHAR (100) DEFAULT '' , "
strSql = strSql & "M_LINK1 VARCHAR (255) DEFAULT '' , "
strSql = strSql & "M_LINK2 VARCHAR (255) DEFAULT '' , "
strSql = strSql & "M_CITY VARCHAR (100) DEFAULT '' , "
strSql = strSql & "M_STATE VARCHAR (100) DEFAULT '' , "
strSql = strSql & "M_PHOTO_URL VARCHAR (255) DEFAULT '', "
strSql = strSql & "PRIMARY KEY (MEMBER_ID), "
strSql = strSql & "KEY " & strMemberTablePrefix & "MEMBERS_MEMBER_ID (MEMBER_ID) ) "

my_Conn.Execute strSql
ChkDBInstall()

strSql = "DROP TABLE IF EXISTS " & strTablePrefix & "MODERATOR "

my_Conn.Execute strSql
ChkDBInstall()

strSql = "CREATE TABLE " & strTablePrefix & "MODERATOR ( "
strSql = strSql & "MOD_ID int (11) DEFAULT '' NOT NULL auto_increment, "
strSql = strSql & "FORUM_ID int (11) DEFAULT '1' , "
strSql = strSql & "MEMBER_ID int (11) DEFAULT '1'  , "
strSql = strSql & "MOD_TYPE smallint (6)  DEFAULT '0', "
strSql = strSql & "PRIMARY KEY (MOD_ID))"
  
my_Conn.Execute strSql
ChkDBInstall()

strSql = "DROP TABLE IF EXISTS " & strTablePrefix & "REPLY "

my_Conn.Execute strSql
ChkDBInstall()

strSql = "CREATE TABLE " & strTablePrefix & "REPLY ( "
strSql = strSql & "CAT_ID int (11) DEFAULT '1' NOT NULL , "
strSql = strSql & "FORUM_ID int (11) DEFAULT '1' NOT NULL , "
strSql = strSql & "TOPIC_ID int (11) DEFAULT '1' NOT NULL , "
strSql = strSql & "REPLY_ID int (11) DEFAULT '' NOT NULL auto_increment, "
strSql = strSql & "R_MAIL smallint (6) DEFAULT '0' , "
strSql = strSql & "R_AUTHOR int (11) DEFAULT '1' , "
strSql = strSql & "R_MESSAGE text , "
strSql = strSql & "R_DATE VARCHAR (50) DEFAULT '' , "
strSql = strSql & "R_IP VARCHAR (50) DEFAULT '000.000.000.000', "
strSql = strSql & "PRIMARY KEY (CAT_ID, FORUM_ID, TOPIC_ID, REPLY_ID), "
strSql = strSql & "KEY " & strTablePrefix & "REPLY_CATFORTOPREPL(CAT_ID,FORUM_ID,TOPIC_ID, REPLY_ID), "
strSql = strSql & "KEY " & strTablePrefix & "REPLY_REP_ID(REPLY_ID), "
strSql = strSql & "KEY " & strTablePrefix & "REPLY_CAT_ID(CAT_ID), "
strSql = strSql & "KEY " & strTablePrefix & "REPLY_FORUM_ID(FORUM_ID), "
strSql = strSql & "KEY " & strTablePrefix & "REPLY_TOPIC_ID (TOPIC_ID) )"

my_Conn.Execute strSql
ChkDBInstall()

strSql = "DROP TABLE IF EXISTS " & strTablePrefix & "TOPICS "

my_Conn.Execute strSql
ChkDBInstall()

strSql = "CREATE TABLE " & strTablePrefix & "TOPICS ( "
strSql = strSql & "CAT_ID int (11) DEFAULT '1' NOT NULL , "
strSql = strSql & "FORUM_ID int (11) DEFAULT '1' NOT NULL , "
strSql = strSql & "TOPIC_ID int (11) DEFAULT '' NOT NULL auto_increment, "
strSql = strSql & "T_STATUS smallint (6) DEFAULT '1' , "
strSql = strSql & "T_MAIL smallint (6) DEFAULT '0' , "
strSql = strSql & "T_SUBJECT VARCHAR (100) DEFAULT '' , "
strSql = strSql & "T_MESSAGE text , "
strSql = strSql & "T_AUTHOR int (11) DEFAULT '1' , "
strSql = strSql & "T_REPLIES int (11) DEFAULT '0' , "
strSql = strSql & "T_VIEW_COUNT int (11) DEFAULT '0' , "
strSql = strSql & "T_LAST_POST VARCHAR (50) DEFAULT '' , "
strSql = strSql & "T_DATE VARCHAR (50) DEFAULT '', "
strSql = strSql & "T_LAST_POSTER int (11) DEFAULT '1', "
strSql = strSql & "T_IP VARCHAR (50) DEFAULT '000.000.000.000', " 
strSql = strSql & "T_LAST_POST_AUTHOR int (11) DEFAULT '1',   "
strSql = strSql & "PRIMARY KEY (CAT_ID, FORUM_ID, TOPIC_ID), "
strSql = strSql & "KEY " & strTablePrefix & "TOPIC_CATFORTOP(CAT_ID,FORUM_ID,TOPIC_ID), "
strSql = strSql & "KEY " & strTablePrefix & "TOPIC_CAT_ID(CAT_ID), "
strSql = strSql & "KEY " & strTablePrefix & "TOPIC_FORUM_ID(FORUM_ID), "
strSql = strSql & "KEY " & strTablePrefix & "TOPIC_TOPIC_ID (TOPIC_ID) )"

my_Conn.Execute strSql
ChkDBInstall()

strSql = "DROP TABLE IF EXISTS " & strTablePrefix & "TOTALS "

my_Conn.Execute strSql
ChkDBInstall()

strSql = "CREATE TABLE " & strTablePrefix & "TOTALS ( "
strSql = strSql & "COUNT_ID smallint (6) DEFAULT '' NOT NULL auto_increment, "
strSql = strSql & "P_COUNT int (11) DEFAULT '0' , "
strSql = strSql & "T_COUNT int (11) DEFAULT '0'  , "
strSql = strSql & "U_COUNT int (11) DEFAULT '0' , "
strSql = strSql & "PRIMARY KEY (COUNT_ID), "
strSql = strSql & "KEY " & strTablePrefix & "TOTALS_COUNT_ID (COUNT_ID) ) "

my_Conn.Execute strSql
ChkDBInstall()

strSql = "DROP TABLE IF EXISTS " & strTablePrefix & "ALLOWED_MEMBERS "

my_Conn.Execute strSql
ChkDBInstall()

strSql = "CREATE TABLE " & strTablePrefix & "ALLOWED_MEMBERS ("
strSql = strSql & "MEMBER_ID INT (11) NOT NULL, FORUM_ID smallint (6) NOT NULL , "
strSql = strSql & "PRIMARY KEY (MEMBER_ID, FORUM_ID) )"

my_Conn.Execute strSql
ChkDBInstall()

strSql = "INSERT INTO " & strTablePrefix & "CONFIG (C_STRVERSION) VALUES('" & strVersion & "') "

my_Conn.Execute strSql
ChkDBInstall()

strSql = "INSERT INTO " & strTablePrefix & "CATEGORY(CAT_STATUS, CAT_NAME) VALUES(1, 'Snitz Forums 2000')"

my_Conn.Execute strSql
ChkDBInstall()

strSql = "INSERT INTO " & strMemberTablePrefix & "MEMBERS (M_STATUS, M_NAME, M_USERNAME, M_PASSWORD, M_EMAIL, M_COUNTRY, "
strSql = strSql & "M_HOMEPAGE, M_SIG, M_DEFAULT_VIEW, M_LEVEL, M_AIM, M_YAHOO, M_ICQ, "
strSql = strSql & "M_POSTS, M_DATE, M_LASTHEREDATE, M_LASTPOSTDATE, M_TITLE, M_SUBSCRIPTION, "
strSql = strSql & "M_HIDE_EMAIL, M_RECEIVE_EMAIL, M_LAST_IP, M_IP) "
strSql = strSql & " VALUES(1, 'Admin', 'Admin', 'admin', 'yourmail@server.com', ' ', ' ', ' ', 1, 3, ' ', ' ', ' ', "
strSql = strSql & " 1, '20001119000000', '20001119000000', '20001119000000', 'Forum Admin', '0', '0', 1, '000.000.000.000', '000.000.000.000')"

my_Conn.Execute strSql
ChkDBInstall()

strSql = "INSERT INTO " & strTablePrefix & "FORUM(CAT_ID, F_STATUS, F_MAIL, F_SUBJECT, F_URL, F_DESCRIPTION, F_TOPICS, F_COUNT, F_LAST_POST, "
strSql = strSql & " F_PASSWORD_NEW, F_USERLIST, F_PRIVATEFORUMS, F_TYPE, F_IP, F_LAST_POST_AUTHOR) "
strSql = strSql & "VALUES(1, 1, '0', 'Testing Forums', '', 'This forum gives you a chance to become more familiar with how this product responds to different features and keeps testing in one place instead of posting tests all over. Happy Posting! <img src=icon_smile.gif border=0 align=middle>', "
strSql = strSql & " 1, 1, '20001119000000', '', '', '0', '0', '000.000.000.000', 1) "

my_Conn.Execute strSql
ChkDBInstall()

strSql = "INSERT INTO " & strTablePrefix & "TOPICS (CAT_ID, FORUM_ID, T_STATUS, T_MAIL, T_SUBJECT, T_MESSAGE, T_AUTHOR, "
strSql = strSql & "T_REPLIES, T_VIEW_COUNT, T_LAST_POST, T_DATE, T_LAST_POSTER, T_IP, T_LAST_POST_AUTHOR) "
strSql = strSql & "VALUES(1, 1, 1, '0', 'Welcome to Snitz Forums 2000', 'Thank you for downloading the Snitz Forums 2000. We hope you enjoy this great tool to support your organization!" & CHR(13) & CHR(10) & CHR(13) & CHR(10) &"Many thanks go out to John Penfold &lt;asp@asp-dev.com&gt; and Tim Teal &lt;tteal@tealnet.com&gt; for the original source code and to all the people of Snitz Forums 2000 at http://forum.snitz.com for continued support of this product.', "
strSql = strSql & "1, '0', '0', '20001119000000', '20001119000000', '0', '000.000.000.000', 1)"

my_Conn.Execute strSql
ChkDBInstall()

strSql = "INSERT INTO " & strTablePrefix & "TOTALS (COUNT_ID, P_COUNT, T_COUNT, U_COUNT) "
strSql = strSql & "VALUES(1,1,1,1)"

my_Conn.Execute strSql
ChkDBInstall()

sub ChkDBInstall()

for counter = 0 to my_conn.Errors.Count -1
	ConnErrorNumber = my_conn.Errors(counter).Number
	ConnErrorDescription = my_conn.Errors(counter).Description
	
	if ConnErrorNumber <> 0 then 
		Err_Msg = "<tr><td bgColor=red align=""left"" width=""30%""><font face=""宋体, Arial, Helvetica"" size=""2""><b>Error: " & ConnErrorNumber & "</b></font></td>"
		Err_Msg = Err_Msg & "<td bgColor=navyblue align=""left""><font face=""宋体, Arial, Helvetica"" size=""2"">" & ConnErrorDescription & "</font></td></tr>"
		Err_Msg = Err_Msg & "<tr><td bgColor=red align=""left"" width=""30%""><font face=""宋体, Arial, Helvetica"" size=""2""><b>strSql: </b></font></td>"
		Err_Msg = Err_Msg & "<td bgColor=navyblue align=""left""><font face=""宋体, Arial, Helvetica"" size=""2"">" & strSql & "</font></td></tr>"	
		
		Response.Write(Err_Msg)
		intCriticalErrors = intCriticalErrors + 1
	end if
next
my_conn.Errors.Clear 

end sub

%>
