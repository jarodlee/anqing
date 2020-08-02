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
Session.LCID = 1033 '## Do Not Edit
Response.Buffer = true

dim strDBType, strConnString, strTablePrefix, strMemberTablePrefix '## Do Not Edit

'#################################################################################
'## SELECT YOUR DATABASE TYPE AND CONNECTION TYPE
'#################################################################################
'strDBType = "sqlserver" '## access, sqlserver or mysql
strDBType = "access" '## access, sqlserver or mysql
'strDBType = "mysql" '## access, sqlserver or mysql

' ## Make sure to uncomment one of the strConnString lines !!
'strConnString = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=c:\inetpub\dbroot\snitz_forums_2000.mdb" '## MS Access 97
'strConnString = "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("tools/snitz2000.mdb")  '## MS Access 97 using virtual path
strConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("tools/snitz2000.mdb") '## MS Access 2000 using virtual path
'strConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=e:\domains\huwr.121host.net\user\data\snitz_forums_2000.mdb;" '## MS Access 2000
'strConnString = "driver={SQL Server};server=192.168.0.2;uid=sa;pwd=wynnston;database=HRForum" '## MS SQL Server 7
'strConnString = "driver=MySQL;server=SERVER_NAME;uid=MYSQL_USER;pwd=PARRWORD;database=DATABASE_NAME" '## MySQL

strTablePrefix = "FORUM_"
strMemberTablePrefix = "FORUM_"

'#################################################################################
'## Do Not Edit Below This Line - It could destroy your forums and lose data
'#################################################################################

dim strVersion, strForumTitle, strCopyright, strTitleImage, strHomeURL, strForumURL, strCookieURL
dim strAuthType, strSetCookieToForum
dim strEmail, strUniqueEmail, strMailMode, strMailServer, strSender
dim strDateType, strTimeAdjust, strTimeType, strForumTimeAdjust
dim strMoveTopicMode, strIPLogging, strPrivateForums, strShowModerators, strShowRank, strAllowForumCode, strAllowHTML
dim strNoCookies, strEditedByDate
dim intHotTopicNum, strHotTopic
dim strIMGInPosts
dim strHomepage, strICQ, strAIM, strInForumAdmin, strSecureAdmin, stricons, strGfxButtons
dim strBadWordFiler, strBadWords
dim strDefaultFontFace, strDefaultFontSize, strHeaderFontSize, strFooterFontSize
dim strPageBGColor, strDefaultFontColor,strPageBGImage
dim strLinkColor, strLinkTextDecoration, strVisitedLinkColor, strVisitedTextDecoration, strActiveLinkColor, strHoverFontColor, strHoverTextDecoration
dim strHeadCellColor, strHeadFontColor, strCategoryCellColor, strCategoryFontColor
dim strForumFirstCellColor, strForumCellColor, strAltForumCellColor, strForumFontColor, strForumLinkColor
dim strTableBorderColor, strPopUpTableColor, strPopUpBorderColor, strNewFontColor, strTopicWidthLeft, strTopicNoWrapLeft, strTopicWidthRight, strTopicNoWrapRight
dim strRankColor1, strRankColor2, strRankColor3
dim strRankLevel0, strRankLevel1, strRankLevel2, strRankLevel3, strRankLevel4, strRankLevel5
dim intRankLevel0, intRankLevel1, intRankLevel2, intRankLevel3, intRankLevel4, intRankLevel5
dim strShowStatistics, strShowImagePoweredBy, strLogonForMail, strShowPaging, strShowTopicNav, strPageSize, strPageNumberSize
dim strNTGroupsSTR
dim strYOffset
dim strStaticYOffset
dim strMenuWidth,strAuthor


strCookieURL = Left(Request.ServerVariables("Path_Info"), InstrRev(Request.ServerVariables("Path_Info"), "/"))
strUniqueID = "Snitz00"

if Application(strCookieURL & "ConfigLoaded")= "" or IsNull(Application(strCookieURL & "ConfigLoaded")) or blnSetup="Y" then 

	'## if the configvariables aren't loaded into the Application object
	'## or after the admin has changed the configuration
	'## the variables get (re)loaded

	set my_Conn = Server.CreateObject("ADODB.Connection")
	my_Conn.Errors.Clear

	my_Conn.Open strConnString

	if blnSetup <> "Y" then 
		for counter = 0 to my_conn.Errors.Count -1
			ConnErrorNumber = my_conn.Errors(counter).Number
			if my_conn.Errors(counter).Number <> 0 then 
				my_conn.Errors.Clear 
				Response.Redirect "setup.asp?RC=1&CC=1&strDBType=" & strDBType & "&EC=" & ConnErrorNumber
			end if
		next
	end if

	my_conn.Errors.Clear 


	'## Forum_SQL
	strSql = "SELECT C_STRVERSION "
	strSql = strSql & ", C_STRFORUMTITLE "
	strSql = strSql & ", C_STRCOPYRIGHT "
	strSql = strSql & ", C_STRTITLEIMAGE "
	strSql = strSql & ", C_STRHOMEURL "
	strSql = strSql & ", C_STRFORUMURL "
	strSql = strSql & ", C_STRAUTHTYPE "
	strSql = strSql & ", C_STRSETCOOKIETOFORUM "
	strSql = strSql & ", C_STREMAIL "
	strSql = strSql & ", C_STRUNIQUEEMAIL "
	strSql = strSql & ", C_STRMAILMODE "
	strSql = strSql & ", C_STRMAILSERVER "
	strSql = strSql & ", C_STRSENDER "
	strSql = strSql & ", C_STRDATETYPE "
	strSql = strSql & ", C_STRTIMEADJUST "
	strSql = strSql & ", C_STRTIMETYPE "
	strSql = strSql & ", C_STRMOVETOPICMODE "
	strSql = strSql & ", C_STRIPLOGGING "
	strSql = strSql & ", C_STRPRIVATEFORUMS "
	strSql = strSql & ", C_STRSHOWMODERATORS "
	strSql = strSql & ", C_STRALLOWFORUMCODE "
	strSql = strSql & ", C_STRALLOWHTML "
	strSql = strSql & ", C_STRNOCOOKIES "
	strSql = strSql & ", C_STRSECUREADMIN "
	strSql = strSql & ", C_STRHOTTOPIC "
	strSql = strSql & ", C_INTHOTTOPICNUM "
	strSql = strSql & ", C_STRIMGINPOSTS "
	strSql = strSql & ", C_STRHOMEPAGE "
	strSql = strSql & ", C_STRICQ "
	strSql = strSql & ", C_STRYAHOO "
	strSql = strSql & ", C_STRAIM "
	strSql = strSql & ", C_stricons "
	strSql = strSql & ", C_STRGFXBUTTONS "
	strSql = strSql & ", C_STREDITEDBYDATE "
	strSql = strSql & ", C_STRBADWORDFILTER "
	strSql = strSql & ", C_STRBADWORDS "
	strSql = strSql & ", C_STRDEFAULTFONTFACE "
	strSql = strSql & ", C_STRDEFAULTFONTSIZE "
	strSql = strSql & ", C_STRHEADERFONTSIZE "
	strSql = strSql & ", C_STRFOOTERFONTSIZE "
	strSql = strSql & ", C_STRPAGEBGCOLOR "
	strSql = strSql & ", C_STRPAGEBGIMAGE "
	strSql = strSql & ", C_STRDEFAULTFONTCOLOR "
	strSql = strSql & ", C_STRLINKCOLOR "
	strSql = strSql & ", C_STRLINKTEXTDECORATION "
	strSql = strSql & ", C_STRVISITEDLINKCOLOR "
	strSql = strSql & ", C_STRVISITEDTEXTDECORATION "
	strSql = strSql & ", C_STRACTIVELINKCOLOR "
	strSql = strSql & ", C_STRHOVERFONTCOLOR "
	strSql = strSql & ", C_STRHOVERTEXTDECORATION "
	strSql = strSql & ", C_STRHEADCELLCOLOR "
	strSql = strSql & ", C_STRHEADFONTCOLOR "
	strSql = strSql & ", C_STRCATEGORYCELLCOLOR "
	strSql = strSql & ", C_STRCATEGORYFONTCOLOR "
	strSql = strSql & ", C_STRFORUMFIRSTCELLCOLOR "
	strSql = strSql & ", C_STRFORUMCELLCOLOR "
	strSql = strSql & ", C_STRALTFORUMCELLCOLOR "
	strSql = strSql & ", C_STRFORUMFONTCOLOR "
	strSql = strSql & ", C_STRFORUMLINKCOLOR "
	strSql = strSql & ", C_STRTABLEBORDERCOLOR "
	strSql = strSql & ", C_STRPOPUPTABLECOLOR "
	strSql = strSql & ", C_STRPOPUPBORDERCOLOR "
	strSql = strSql & ", C_STRNEWFONTCOLOR "
	strSql = strSql & ", C_STRTOPICWIDTHLEFT "
	strSql = strSql & ", C_STRTOPICNOWRAPLEFT "
	strSql = strSql & ", C_STRTOPICWIDTHRIGHT "
	strSql = strSql & ", C_STRTOPICNOWRAPRIGHT "
	strSql = strSql & ", C_STRSHOWRANK "
	strSql = strSql & ", C_STRRANKADMIN "
	strSql = strSql & ", C_STRRANKMOD "
	strSql = strSql & ", C_STRRANKLEVEL0 "
	strSql = strSql & ", C_STRRANKLEVEL1 "
	strSql = strSql & ", C_STRRANKLEVEL2 "
	strSql = strSql & ", C_STRRANKLEVEL3 "
	strSql = strSql & ", C_STRRANKLEVEL4 "
	strSql = strSql & ", C_STRRANKLEVEL5 "
	strSql = strSql & ", C_STRRANKCOLORADMIN "
	strSql = strSql & ", C_STRRANKCOLORMOD "
	strSql = strSql & ", C_STRRANKCOLOR0 "
	strSql = strSql & ", C_STRRANKCOLOR1 "
	strSql = strSql & ", C_STRRANKCOLOR2 "
	strSql = strSql & ", C_STRRANKCOLOR3 "
	strSql = strSql & ", C_STRRANKCOLOR4 "
	strSql = strSql & ", C_STRRANKCOLOR5 "
	strSql = strSql & ", C_INTRANKLEVEL0 "
	strSql = strSql & ", C_INTRANKLEVEL1 "
	strSql = strSql & ", C_INTRANKLEVEL2 "
	strSql = strSql & ", C_INTRANKLEVEL3 "
	strSql = strSql & ", C_INTRANKLEVEL4 "
	strSql = strSql & ", C_INTRANKLEVEL5 "
	strSql = strSql & ", C_STRSIGNATURES "
	strSql = strSql & ", C_STRSHOWSTATISTICS "
	strSql = strSql & ", C_STRSHOWIMAGEPOWEREDBY "
	strSql = strSql & ", C_STRLOGONFORMAIL "
	strSql = strSql & ", C_STRSHOWPAGING "
	strSql = strSql & ", C_STRSHOWTOPICNAV "
	strSql = strSql & ", C_STRPAGESIZE "
	strSql = strSql & ", C_STRPAGENUMBERSIZE "
	strSql = strSql & ", C_STRFULLNAME"
	strSql = strSql & ", C_STRPICTURE"
	strSql = strSql & ", C_STRSEX"
	strSql = strSql & ", C_STRCITY"
	strSql = strSql & ", C_STRSTATE"
	strSql = strSql & ", C_STRAGE"
	strSql = strSql & ", C_STRCOUNTRY"
	strSql = strSql & ", C_STROCCUPATION"
	strSql = strSql & ", C_STRBIO"
	strSql = strSql & ", C_STRHOBBIES"
	strSql = strSql & ", C_STRLNEWS"
	strSql = strSql & ", C_STRQUOTE"
	strSql = strSql & ", C_STRMARSTATUS"
	strSql = strSql & ", C_STRFAVLINKS"
	strSql = strSql & ", C_STRRECENTTOPICS"
	strSql = strSql & ", C_STRHOMEPAGE"
	strSql = strSql & ", C_STRNTGROUPS"
	strSql = strSql & ", C_STRAUTOLOGON"
	strSql = strSql & " FROM " & strTablePrefix & "CONFIG "

	'Response.Write strSQL
	'Response.End

	set rsConfig = my_Conn.Execute (strSql)
	
'############ HuwR Config stuff #####################
	Dim modDict
	Set modDict = Server.CreateObject("Scripting.Dictionary")
	set objRec = my_Conn.execute("SELECT M_CODE, M_VALUE FROM " & strTablePrefix & "MODS WHERE M_NAME = 'HModEnable'")

	while not objRec.EOF 
	modDict.Add objRec.Fields.Item("m_code").Value, objRec.Fields.Item("m_value").Value
	objRec.moveNext
	wend 
	intPMessages = modDict.Item("Pmessages")
	intUserFields = modDict.Item("UserFields")
	intAllowUploads = modDict.Item("Attachment")
	intPollMentor = modDict.Item("PollMentor")
	intSideMenu = modDict.Item("SideMenu")
	strImageUrl = modDict.Item("imageURLPath")
	intNewMemberPM = modDict.Item("NewMemberPM")
	
	on error resume next
	set sc = GetObject("java:SpellCheck")
	if err.number <> 0 then
		intSpellCheck = 0
	else
		intSpellCheck = 1
	end if
	set sc = Nothing

'############ HuwR Config stuff #####################

	Application.Lock
	Application(strCookieURL & "strVersion") = rsConfig("C_STRVERSION")
	Application(strCookieURL & "strForumTitle") = rsConfig("C_STRFORUMTITLE")
	Application(strCookieURL & "strCopyright") = rsConfig("C_STRCOPYRIGHT")
	Application(strCookieURL & "strTitleImage") = rsConfig("C_STRTITLEIMAGE")
	Application(strCookieURL & "strHomeURL") = rsConfig("C_STRHOMEURL")
	Application(strCookieURL & "strForumURL") = rsConfig("C_STRFORUMURL")

	Application(strCookieURL & "strAuthType") = rsConfig("C_STRAUTHTYPE")
	Application(strCookieURL & "strSetCookieToForum") = rsConfig("C_STRSETCOOKIETOFORUM")
	Application(strCookieURL & "strEmail") = rsConfig("C_STREMAIL")
	Application(strCookieURL & "strUniqueEmail") = rsConfig("C_STRUNIQUEEMAIL")
	Application(strCookieURL & "strMailMode") = rsConfig("C_STRMAILMODE")
	Application(strCookieURL & "strMailServer") = rsConfig("C_STRMAILSERVER")
	Application(strCookieURL & "strSender") = rsConfig("C_STRSENDER")
	Application(strCookieURL & "strDateType") = rsConfig("C_STRDATETYPE")
	Application(strCookieURL & "strTimeAdjust") = rsConfig("C_STRTIMEADJUST")
	Application(strCookieURL & "strTimeType") = rsConfig("C_STRTIMETYPE")
	Application(strCookieURL & "strMoveTopicMode") = rsConfig("C_STRMOVETOPICMODE")
	Application(strCookieURL & "strIPLogging") = rsConfig("C_STRIPLOGGING")
	Application(strCookieURL & "strPrivateForums") = rsConfig("C_STRPRIVATEFORUMS")
	Application(strCookieURL & "strShowModerators") = rsConfig("C_STRSHOWMODERATORS")
	Application(strCookieURL & "strAllowForumCode") = rsConfig("C_STRALLOWFORUMCODE")
	Application(strCookieURL & "strIMGInPosts") = rsConfig("C_STRIMGINPOSTS")
	Application(strCookieURL & "strAllowHTML") = rsConfig("C_STRALLOWHTML")
	Application(strCookieURL & "strNoCookies") = rsConfig("C_STRNOCOOKIES")
	Application(strCookieURL & "strSecureAdmin") = rsConfig("C_STRSECUREADMIN")
	Application(strCookieURL & "strHotTopic") = rsConfig("C_STRHOTTOPIC")
	Application(strCookieURL & "intHotTopicNum") = rsConfig("C_INTHOTTOPICNUM")
	Application(strCookieURL & "strHomepage") = rsConfig("C_STRHOMEPAGE")
	Application(strCookieURL & "strICQ") = rsConfig("C_STRICQ")
	Application(strCookieURL & "strYAHOO") = rsConfig("C_STRYAHOO")
	Application(strCookieURL & "strAIM") = rsConfig("C_STRAIM")
	Application(strCookieURL & "stricons") = rsConfig("C_stricons")
	Application(strCookieURL & "strGfxButtons") = rsConfig("C_STRGFXBUTTONS")
	Application(strCookieURL & "strEditedByDate") = rsConfig("C_STREDITEDBYDATE")
	Application(strCookieURL & "strBadWordFilter") = rsConfig("C_STRBADWORDFILTER")
	Application(strCookieURL & "strBadWords") = rsConfig("C_STRBADWORDS")
	Application(strCookieURL & "strDefaultFontFace") = rsConfig("C_STRDEFAULTFONTFACE")
	Application(strCookieURL & "strDefaultFontSize") = rsConfig("C_STRDEFAULTFONTSIZE")
	Application(strCookieURL & "strHeaderFontSize") = rsConfig("C_STRHEADERFONTSIZE")
	Application(strCookieURL & "strFooterFontSize") = rsConfig("C_STRFOOTERFONTSIZE")
	Application(strCookieURL & "strPageBGColor") = rsConfig("C_STRPAGEBGCOLOR")
	Application(strCookieURL & "strPageBGImage") = rsConfig("C_STRPAGEBGIMAGE")
	Application(strCookieURL & "strDefaultFontColor") = rsConfig("C_STRDEFAULTFONTCOLOR")
	Application(strCookieURL & "strLinkColor") = rsConfig("C_STRLINKCOLOR")
	Application(strCookieURL & "strLinkTextDecoration") = rsConfig("C_STRLINKTEXTDECORATION")
	Application(strCookieURL & "strVisitedLinkColor") = rsConfig("C_STRVISITEDLINKCOLOR")
	Application(strCookieURL & "strVisitedTextDecoration") = rsConfig("C_STRVISITEDTEXTDECORATION")
	Application(strCookieURL & "strActiveLinkColor") = rsConfig("C_STRACTIVELINKCOLOR")
	Application(strCookieURL & "strHoverFontColor") = rsConfig("C_STRHOVERFONTCOLOR")
	Application(strCookieURL & "strHoverTextDecoration") = rsConfig("C_STRHOVERTEXTDECORATION")
	Application(strCookieURL & "strHeadCellColor") = rsConfig("C_STRHEADCELLCOLOR")
	Application(strCookieURL & "strHeadFontColor") = rsConfig("C_STRHEADFONTCOLOR")
	Application(strCookieURL & "strCategoryCellColor") = rsConfig("C_STRCATEGORYCELLCOLOR")
	Application(strCookieURL & "strCategoryFontColor") = rsConfig("C_STRCATEGORYFONTCOLOR")
	Application(strCookieURL & "strForumFirstCellColor") = rsConfig("C_STRFORUMFIRSTCELLCOLOR")
	Application(strCookieURL & "strForumCellColor") = rsConfig("C_STRFORUMCELLCOLOR")
	Application(strCookieURL & "strAltForumCellColor") = rsConfig("C_STRALTFORUMCELLCOLOR")
	Application(strCookieURL & "strForumFontColor") = rsConfig("C_STRFORUMFONTCOLOR")
	Application(strCookieURL & "strForumLinkColor") = rsConfig("C_STRFORUMLINKCOLOR")
	Application(strCookieURL & "strTableBorderColor") = rsConfig("C_STRTABLEBORDERCOLOR")
	Application(strCookieURL & "strPopUpTableColor") = rsConfig("C_STRPOPUPTABLECOLOR")
	Application(strCookieURL & "strPopUpBorderColor") = rsConfig("C_STRPOPUPBORDERCOLOR")
	Application(strCookieURL & "strNewFontColor") = rsConfig("C_STRNEWFONTCOLOR")
	Application(strCookieURL & "strTopicWidthLeft") = rsConfig("C_STRTOPICWIDTHLEFT")
	Application(strCookieURL & "strTopicNoWrapLeft") = rsConfig("C_STRTOPICNOWRAPLEFT")
	Application(strCookieURL & "strTopicWidthRight") = rsConfig("C_STRTOPICWIDTHRIGHT")
	Application(strCookieURL & "strTopicNoWrapRight") = rsConfig("C_STRTOPICNOWRAPRIGHT")
	Application(strCookieURL & "strShowRank") = rsConfig("C_STRSHOWRANK")
	Application(strCookieURL & "strRankAdmin") = rsConfig("C_STRRANKADMIN")
	Application(strCookieURL & "strRankMod") = rsConfig("C_STRRANKMOD")
	Application(strCookieURL & "strRankLevel0") = rsConfig("C_STRRANKLEVEL0")
	Application(strCookieURL & "strRankLevel1") = rsConfig("C_STRRANKLEVEL1")
	Application(strCookieURL & "strRankLevel2") = rsConfig("C_STRRANKLEVEL2")
	Application(strCookieURL & "strRankLevel3") = rsConfig("C_STRRANKLEVEL3")
	Application(strCookieURL & "strRankLevel4") = rsConfig("C_STRRANKLEVEL4")
	Application(strCookieURL & "strRankLevel5") = rsConfig("C_STRRANKLEVEL5")
	Application(strCookieURL & "strRankColorAdmin") = rsConfig("C_STRRANKCOLORADMIN")
	Application(strCookieURL & "strRankColorMod") = rsConfig("C_STRRANKCOLORMOD")
	Application(strCookieURL & "strRankColor0") = rsConfig("C_STRRANKCOLOR0")
	Application(strCookieURL & "strRankColor1") = rsConfig("C_STRRANKCOLOR1")
	Application(strCookieURL & "strRankColor2") = rsConfig("C_STRRANKCOLOR2")
	Application(strCookieURL & "strRankColor3") = rsConfig("C_STRRANKCOLOR3")
	Application(strCookieURL & "strRankColor4") = rsConfig("C_STRRANKCOLOR4")
	Application(strCookieURL & "strRankColor5") = rsConfig("C_STRRANKCOLOR5")
	Application(strCookieURL & "intRankLevel0") = rsConfig("C_INTRANKLEVEL0")
	Application(strCookieURL & "intRankLevel1") = rsConfig("C_INTRANKLEVEL1")
	Application(strCookieURL & "intRankLevel2") = rsConfig("C_INTRANKLEVEL2")
	Application(strCookieURL & "intRankLevel3") = rsConfig("C_INTRANKLEVEL3")
	Application(strCookieURL & "intRankLevel4") = rsConfig("C_INTRANKLEVEL4")
	Application(strCookieURL & "intRankLevel5") = rsConfig("C_INTRANKLEVEL5")
	Application(strCookieURL & "strShowStatistics") = rsconfig("C_STRSHOWSTATISTICS")
	Application(strCookieURL & "strShowImagePoweredBy") = rsconfig("C_STRSHOWIMAGEPOWEREDBY")
	Application(strCookieURL & "strLogonForMail") = rsconfig("C_STRLOGONFORMAIL")
	Application(strCookieURL & "strShowPaging") = rsconfig("C_STRSHOWPAGING")
	Application(strCookieURL & "strShowTopicNav") = rsconfig("C_STRSHOWTOPICNAV")
	Application(strCookieURL & "strPageSize") = rsconfig("C_STRPAGESIZE")
	Application(strCookieURL & "strPageNumberSize") = rsconfig("C_STRPAGENUMBERSIZE")
	Application(strCookieURL & "strFullName") = rsconfig("C_STRFULLNAME")
	Application(strCookieURL & "strPicture") = rsconfig("C_STRPICTURE")
	Application(strCookieURL & "strSex") = rsconfig("C_STRSEX")
	Application(strCookieURL & "strCity") = rsconfig("C_STRCITY")
	Application(strCookieURL & "strState") = rsconfig("C_STRSTATE")
	Application(strCookieURL & "strAge") = rsconfig("C_STRAGE")
	Application(strCookieURL & "strCountry") = rsconfig("C_STRCOUNTRY")
	Application(strCookieURL & "strOccupation") = rsconfig("C_STROCCUPATION")
	Application(strCookieURL & "strBio") = rsconfig("C_STRBIO")
	Application(strCookieURL & "strHobbies") = rsconfig("C_STRHOBBIES")
	Application(strCookieURL & "strLNews") = rsconfig("C_STRLNEWS")
	Application(strCookieURL & "strQuote") = rsconfig("C_STRQUOTE")
	Application(strCookieURL & "strMarStatus") = rsconfig("C_STRMARSTATUS")
	Application(strCookieURL & "strFavLinks") = rsconfig("C_STRFAVLINKS")
	Application(strCookieURL & "strRecentTopics") = rsconfig("C_STRRECENTTOPICS")
	Application(strCookieURL & "strHomePage") = rsconfig("C_STRHOMEPAGE")
	Application(strCookieURL & "STRNTGROUPS") = rsConfig("C_STRNTGROUPS")
	Application(strCookieURL & "STRAUTOLOGON") = rsConfig("C_STRAUTOLOGON")
'Rem User Field Code #######################################
	Application(strCookieURL & "intUserFields") = modDict.Item("UserFields")
	Application(strCookieURL & "intPollMentor") = modDict.Item("PollMentor")
	Application(strCookieURL & "intAllowUploads") = modDict.Item("Attachment")
	Application(strCookieURL & "intSideMenu") = modDict.Item("SideMenu")
	Application(strCookieURL & "intPMessages") = modDict.Item("Pmessages")
	Application(strCookieURL & "strImagePath") = modDict.Item("imageURLPath")
	Application(strCookieURL & "intSpellcheck") = intSpellcheck
	Application(strCookieURL & "intNewMemberPM") = modDict.Item("NewMemberPM")
'Rem User Field Code #######################################
	Application.UnLock

	if blnSetup <> "Y" then 
		for counter = 0 to my_Conn.Errors.Count -1
			if my_Conn.Errors(counter).Number <> 0 or Err.number > 0 then 
		 		myConnError = my_Conn.Errors(counter).Number
				my_Conn.Errors.Clear 
				Err.Clear
		
				strSql = "SELECT " & strTablePrefix & "CONFIG.C_STRVERSION, "
				strSql = strSql & strTablePrefix & "CONFIG.C_STRSENDER "
				strSql = strSql & " FROM " & strTablePrefix & "CONFIG "

				set rsInfo = my_Conn.Execute (StrSql)
				strVersion = rsInfo("C_STRVERSION")
				strSender = rsInfo("C_STRSENDER")
		
				rsInfo.Close
				set rsInfo = nothing
				my_Conn.Close
				set my_Conn = nothing
		
				Response.Redirect "setup.asp?RC=2&CC=1&MAIL=" & Server.UrlEncode(strSender) & "&VER=" & Server.URLEncode(strVersion) & "&strDBType="& strDBType & "&EC=" & myConnError
			end if
		next
	end if

	my_Conn.Close
	set my_Conn = nothing
	Application.Lock
	Application(strCookieURL & "ConfigLoaded")= "YES"
	Application.UnLock
end if 

' ## Read the config-info from the application variables...

strVersion = Application(strCookieURL & "strVersion")
strForumTitle = Application(strCookieURL & "strForumTitle")
strCopyright = Application(strCookieURL & "strCopyright")
strTitleImage = Application(strCookieURL & "strTitleImage")
strHomeURL = Application(strCookieURL & "strHomeURL")
strForumURL = Application(strCookieURL & "strForumURL")
strAuthType = Application(strCookieURL & "strAuthType")
strSetCookieToForum = Application(strCookieURL & "strSetCookieToForum")
strEmail = Application(strCookieURL & "strEmail")
strUniqueEmail = Application(strCookieURL & "strUniqueEmail")
strMailMode = Application(strCookieURL & "strMailMode")
strMailServer = Application(strCookieURL & "strMailServer")
strSender = Application(strCookieURL & "strSender")
strDateType = Application(strCookieURL & "strDateType")
strTimeAdjust = Application(strCookieURL & "strTimeAdjust")
strTimeType = Application(strCookieURL & "strTimeType")
strMoveTopicMode = Application(strCookieURL & "strMoveTopicMode")
strIPLogging = Application(strCookieURL & "strIPLogging")
strPrivateForums = Application(strCookieURL & "strPrivateForums")
strShowModerators = Application(strCookieURL & "strShowModerators")
strAllowForumCode = Application(strCookieURL & "strAllowForumCode")
strIMGInPosts = Application(strCookieURL & "strIMGInPosts")
strAllowHTML = Application(strCookieURL & "strAllowHTML")
strNoCookies = Application(strCookieURL & "strNoCookies")
strSecureAdmin = Application(strCookieURL & "strSecureAdmin")
strHotTopic = Application(strCookieURL & "strHotTopic")
intHotTopicNum = Application(strCookieURL & "intHotTopicNum")
strICQ = Application(strCookieURL & "strICQ")
strYAHOO = Application(strCookieURL & "strYAHOO")
strAIM = Application(strCookieURL & "strAIM")
strFullName = Application(strCookieURL & "strFullName")
strPicture = Application(strCookieURL & "strPicture")
strSex = Application(strCookieURL & "strSex")
strCity= Application(strCookieURL & "strCity")
strState = Application(strCookieURL & "strState")
strAge = Application(strCookieURL & "strAge")
strCountry = Application(strCookieURL & "strCountry") 
strOccupation = Application(strCookieURL & "strOccupation")
strBio = Application(strCookieURL & "strBio") 
strHobbies = Application(strCookieURL & "strHobbies") 
strLNews = 	Application(strCookieURL & "strLNews") 
strQuote = Application(strCookieURL & "strQuote") 
strMarStatus = Application(strCookieURL & "strMarStatus")
strFavLinks = Application(strCookieURL & "strFavLinks")
strRecentTopics = Application(strCookieURL & "strRecentTopics") 
strAllowHideEmail = "1" '##not yet used !
strHomepage = Application(strCookieURL & "strHomepage")
strUseExtendedProfile = (strBio + strHobbies + strLNews + strRecentTopics + strPicture) > 0
strUseExtendedProfile = strUseExtendedProfile or ((strICQ + strYAHOO + strAIM + strFullName*2 + strSex + strCity + strState + strAge + strCountry + strOccupation + strFavLinks*2) > 5)
'Rem User Field Code #######################################
intPollMentor = Application(strCookieURL & "intPollMentor")
intAllowUploads = Application(strCookieURL & "intAllowUploads")
intSideMenu = Application(strCookieURL & "intSideMenu")
intPMessages = Application(strCookieURL & "intPMessages")
intSpellcheck = Application(strCookieURL & "intSpellcheck")
intUserFields = Application(strCookieURL & "intUserFields")
strImageURL = Application(strCookieURL & "strImagePath")
intNewMemberPM = Application(strCookieURL & "intNewMemberPM")
strUseExtendedProfile = strUseExtendedProfile or ((intUserFields ) > 0)
'Rem User Field Code #######################################
stricons = Application(strCookieURL & "stricons")
strGfxButtons = Application(strCookieURL & "strGfxButtons")
strEditedByDate = Application(strCookieURL & "strEditedByDate")
strBadWordFilter = Application(strCookieURL & "strBadWordFilter")
strBadWords = Application(strCookieURL & "strBadWords")
strDefaultFontFace = Application(strCookieURL & "strDefaultFontFace")
strDefaultFontSize = Application(strCookieURL & "strDefaultFontSize")
strHeaderFontSize = Application(strCookieURL & "strHeaderFontSize")
strFooterFontSize = Application(strCookieURL & "strFooterFontSize")
strPageBGColor = Application(strCookieURL & "strPageBGColor")
strPageBGImage = Application(strCookieURL & "strPageBGImage")
strDefaultFontColor = Application(strCookieURL & "strDefaultFontColor")
strLinkColor = Application(strCookieURL & "strLinkColor")
strLinkTextDecoration = Application(strCookieURL & "strLinkTextDecoration")
strVisitedLinkColor = Application(strCookieURL & "strVisitedLinkColor")
strVisitedTextDecoration = Application(strCookieURL & "strVisitedTextDecoration")
strActiveLinkColor = Application(strCookieURL & "strActiveLinkColor")
strHoverFontColor = Application(strCookieURL & "strHoverFontColor")
strHoverTextDecoration = Application(strCookieURL & "strHoverTextDecoration")
strHeadCellColor = Application(strCookieURL & "strHeadCellColor")
strHeadFontColor = Application(strCookieURL & "strHeadFontColor")
strCategoryCellColor = Application(strCookieURL & "strCategoryCellColor")
strCategoryFontColor = Application(strCookieURL & "strCategoryFontColor")
strForumFirstCellColor = Application(strCookieURL & "strForumFirstCellColor")
strForumCellColor = Application(strCookieURL & "strForumCellColor")
strAltForumCellColor = Application(strCookieURL & "strAltForumCellColor")
strForumFontColor = Application(strCookieURL & "strForumFontColor")
strForumLinkColor = Application(strCookieURL & "strForumLinkColor")
strTableBorderColor = Application(strCookieURL & "strTableBorderColor")
strPopUpTableColor = Application(strCookieURL & "strPopUpTableColor")
strPopUpBorderColor = Application(strCookieURL & "strPopUpBorderColor")
strNewFontColor = Application(strCookieURL & "strNewFontColor")
strTopicWidthLeft = Application(strCookieURL & "strTopicWidthLeft")
strTopicNoWrapLeft = Application(strCookieURL & "strTopicNoWrapLeft")
strTopicWidthRight = Application(strCookieURL & "strTopicWidthRight")
strTopicNoWrapRight = Application(strCookieURL & "strTopicNoWrapRight")
strShowRank = Application(strCookieURL & "strShowRank")
strRankAdmin = Application(strCookieURL & "strRankAdmin")
strRankMod = Application(strCookieURL & "strRankMod")
strRankLevel0 = Application(strCookieURL & "strRankLevel0")
strRankLevel1 = Application(strCookieURL & "strRankLevel1")
strRankLevel2 = Application(strCookieURL & "strRankLevel2")
strRankLevel3 = Application(strCookieURL & "strRankLevel3")
strRankLevel4 = Application(strCookieURL & "strRankLevel4")
strRankLevel5 = Application(strCookieURL & "strRankLevel5")
strRankColorAdmin = Application(strCookieURL & "strRankColorAdmin")
strRankColorMod = Application(strCookieURL & "strRankColorMod")
strRankColor0 = Application(strCookieURL & "strRankColor0")
strRankColor1 = Application(strCookieURL & "strRankColor1")
strRankColor2 = Application(strCookieURL & "strRankColor2")
strRankColor3 = Application(strCookieURL & "strRankColor3")
strRankColor4 = Application(strCookieURL & "strRankColor4")
strRankColor5 = Application(strCookieURL & "strRankColor5")
intRankLevel0 = Application(strCookieURL & "intRankLevel0")
intRankLevel1 = Application(strCookieURL & "intRankLevel1")
intRankLevel2 = Application(strCookieURL & "intRankLevel2")
intRankLevel3 = Application(strCookieURL & "intRankLevel3")
intRankLevel4 = Application(strCookieURL & "intRankLevel4")
intRankLevel5 = Application(strCookieURL & "intRankLevel5")
strShowStatistics = Application(strCookieURL & "strShowStatistics")
strShowImagePoweredBy = Application(strCookieURL & "strShowImagePoweredBy")
strLogonForMail = Application(strCookieURL & "strLogonForMail")
strShowPaging = Application(strCookieURL & "strShowPaging")
strShowTopicNav = Application(strCookieURL & "strShowTopicNav")
strPageSize = Application(strCookieURL & "strPageSize")
strPageNumberSize = Application(strCookieURL & "strPageNumberSize")
strForumTimeAdjust = DateAdd("h", strTimeAdjust , Now())
strNTGroups = Application(strCookieURL & "STRNTGROUPS")
strAutoLogon = Application(strCookieURL & "STRAUTOLOGON")

if strSecureAdmin = "0" then
	Session(strCookieURL & "Approval") = "15916941253"
end if

on error goto 0

if strAuthType = "db" then
	strDBNTSQLName = "M_NAME"
	strAutoLogon ="0"
	strNTGroups  ="0"
else
	strDBNTSQLName = "M_USERNAME"
end if
%>
