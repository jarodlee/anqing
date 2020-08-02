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
<!--#INCLUDE FILE="inc_top_short.asp" -->

<table border="0" width="95%" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td bgcolor="<% =strTableBorderColor %>">
<table border="0" width="100%" cellspacing="1" cellpadding="4">
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="strConnString"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>How do I configure the strConnString?</b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>">
    <li><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>"><b>DSN:</b></font><br>
    <font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">snitz_forum</font></li>
    <li><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>"><b>MS Access DSN-less:</b></font><br>
    <font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">strConnString = &quot;DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=c:\www\snitz.com\db\snitz_forum.mdb</font></li>
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="tableprefix"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>What's Table Name Prefix?</b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    Table Name Prefix is used if you have multiple versions of the forum running in the same database. This way you can name the tables differently and still use one user to connect. (eg. FORUM_ and FORUM2_)</FONT>
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="FORUMtitle"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    What's Forum Title?
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    Forum Title is the title that shows up in the upper right hand corner of the forum. It is also used in email's to show where the email came from when posting replies are sent and when new users register.
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="copyright"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    What's Forum Copyright?
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    This copyright statements location is basicaly saying that any topics or replies that are posted are copyrighted material of your organization. This copyright location also helps to copyright the images of your logo and any other material that may be posted on forum pages; however, it is understood by copyright statements in code and informational pages, that the forum code itself is still copyright &copy; 2000 Snitz Communications.
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="titleimage"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    What's Title Image?
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    Use a relative path to point to the image you want to show up in the upper left-hand corner of your forum window.<br>
    <br>
    For example:<br>
    <b>bboard_snitz.gif</b><br>
    This points to the bboard_snitz.gif graphic in the same directory, whereas the following would point to the root of the web server and up into the base /<%=strImageURL %> directory:<br>
    <b>../<%=strImageURL %>bboard_snitz.gif</b>
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="homeurl"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    What's the Home URL?
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    The Home URL is the base address for your website. An example would be:<br>
    <b>forum.snitz.com</b><br>
    <br>
    NOTE: Include the full path of the URL whether that begin with <b>http://</b> in front or a relative URL such as <b>../</b>.
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="forumurl"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    What's the Forum URL?
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    The Forum URL is the base address for your forum. An example would be:<br>
    <b>forum.snitz.com/forum</b>
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="AuthType"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    Authorization Type?
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    You can either select DataBase or NT Domain authorization.
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="SetCookieToForum"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    Set Cookie To...
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    You can tell your forum to set it's cookie to either the forum, or the base website. You would set it to the forum if you were hosting multiple forums on the same server or the same domain, and they each had different user communities, otherwise you want this feature set to Website and NOT Forum.
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="GfxButtons"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    Graphic Buttons?
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    By enabling this feature, the forums will use pictures/graphics instead of the default buttons for "Submit" and "Reset" etc...
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="timetype"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    Time Display?
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    Choose 24Hr to display all times in military (24 hour) format or 12Hr to display all times in 12 hour format appended with an AM or PM depending on whether it's before or after midday. Default is 24 hour.
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr> 
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="datetype"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    Date Display?
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    Choose the format you wish all dates to be displayed in. Default is 12/31/2000 (US Short).
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr> 
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="TimeAdjust"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    Time Adjustment?
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    Enter either a positive or negative integer value between +12 and 0 and -12. This may come in handy if your located in one part of the world, and your server is in another, and you need the time displayed in the forum to be converted to a local time for you! (Default value is 0, meaning no adjustment)
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr> 
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="email"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    What does Email do?
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    Disabling the Email function will turn off any features that involve sending mail. If you don't have an SMTP server of any type, you will want to turn this feature off. If you do have an SMTP (mail) server, however, then also select the type of server you have from the dropdown menu.
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="mailserver"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    What is a Mail Server Address?
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    The mail server address is the actual domain name that resolves your mail server. This could be something like:<br>
    <b>mail.snitz.com</b><br>
    or it could be the same address as the web server:<br>
    <b>www.snitz.com</b><br>
    Either way, don't put the <b>http://</b> on it.<br>
    <br>
    <b>注意：</b> If you are using CDONTS as a mail server type, you do not need to fill in this field.
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="sender"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    Administrator Email Address?
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    This address is referenced by the forums in a couple ways.<br>
    <ol>
    <li>When mail is sent, it is sent from this user email account.</li>
    <li>This user is also the point of contact given if there is a problem with these forums.</li>
    </ol>
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="UniqueEmail"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    Unique Email Address?
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    Do you want to require each user to have their own email address?
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="LogonForMail"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    Require Logon for sending Mail?
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    Do you require a user to be logged on before being able to use the <i>Send Topic To a Friend</i> or <i>Email Poster</i> options ?
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="privateforums"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    What are Private Forums?
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    Private Forums enable you to only allow certain members to see that the forum exists. If it's only password protected, everyone can see that it exists, however, they are prompted for a password to get in.
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="MoveTopicMode"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    Why Restrict Moderators from Moving Posts?
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    This feature either allows or dis-allows a Moderator of one forum to move topics within their forum to someone else's forum where they do not have moderator rights.
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="ShowRank"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    Showing Ranks?
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    <ol>
      <li>Don't Show Any</li>
      <li>Show Rank Only</li>
      <li>Show Stars Only</li>
      <li>Show Both Stars and Rank</li>
    </ol>
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="IPLogging"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    What is IP Logging?
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    IP Logging will record in the database the IP address of the person who posted a new Topic or Reply. A moderator or administrator then could view the IP by clicking on an <%=strImageURL %>icon above the post in the topic.
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="ShowModerator"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    What does Show Moderators do?
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    Basically, if this function is on, it shows the name of the moderator beside the forum that they moderate on the main default page. If it is off, however, visitors won't see who is moderating the forum they are posting in.
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="AllowForumCode"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    Enable Forum Code?
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    By turning off Forum Code, you can allow users to mark up their posts with safe codes.
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="AllowHTML"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    Why would I allow HTML?
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    By allowing HTML you are opening up a whole big can of worms. You may wish to allow HTML in a controlled INTRANET environment,though. It is not recommended to be used on the INTERNET as anyone can post anything without your being able to screen it. IE Pornographic pictures, JavaScript that messes up your pages, etc...
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="hottopics"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    What are Hot Topics?
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    Hot Topics change the topic folder <%=strImageURL %>icon in the Forum view from a normal folder to a flaming folder to let people know that your minimum number of posts has been met to categorize this topic as one that is seeing a lot of action.
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="imginposts"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    Why enable Images in Posts?
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    Allows users to place images into their Posts. However, you should be aware that this feature would allow anyone to post ANY image in your forums. This may lead to broken links and potentially objectionable material being displayed.
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="homepages"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    What is Homepages For?
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    Allow your users to display their homepage link by their name on each post.
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="icq"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    What is the ICQ Option For?
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    Turns On/Off features that allow users to enter their ICQ number... then for other users to send them ICQ messages and/or see if they are online.
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="yahoo"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    What is the YAHOO Option For?
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    Turns On/Off features that allow users to enter their YAHOO username... then for other users to send them messages and/or add them to their buddy list.
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="aim"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    What is the AIM Option For?
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    Turns On/Off features that allow users to enter their AIM username... then for other users to send them messages and/or add them to their buddy list.
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="<%=strImageURL %>icons"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    What do <%=strImageURL %>icons do?
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    Allow users to post smiley faces and other <%=strImageURL %>icons allowed by the Forums within the body of their posts!
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="stats"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    What does Detailed Statistics do?
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    Turns On/Off the display of detailed statistics (last visited date and time, last post, active topics, newest member) at the bottom of the forum. 
    When turned off, some statistics are displayed at the top of the page.
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="showadministration"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    Show Administration?
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    <b>注意：</b> This feature will go away by the time version 4.0 comes out and will utilize cookies (or some similar method) to determine what a user may or may not see.
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="secureadminmode"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    Secure Admin Mode?
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>" color="red">
    <b>WARNING: Only turn Secure Admin off if you absolutely need to. If this option is turned off, anyone can change your forum's configuration!</b>
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="allownoncookies"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    Why would I want Non-Cookie Mode on?
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    If your user base does not use cookies, then you would want to turn this function "ON". WARNING: all your admin functions will be visible to all users if this function is "ON".
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="RankColor"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    Color of Stars?
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    You can change the color of stars that show up for each rank of member. (only when the Stars function is turned on)
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="editedbydate"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    What would Edited By on Date do?
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    When a post is edited, there is an appending to the end of the post that says when and by whom the post was edited. Turning this function off would make it so that the footer would not be placed on the end of the post.
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="badwordfilter"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    Bad Word Filter?
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    Screen out words you and your guests would find offensive.
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="columnwidth"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    How does Column Width Work?
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    This sets the width of the column in question. It is not recommended that you change this unless you really know what your doing.
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="nowrap"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    What is NOWRAP?
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    NOWRAP prevents the text in a column from auto wrapping. This could be bad if you have people posting long strings of text in the right column (message box), reason being: this would cause an awful long horizontal scroll bar in most cases.
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="fontfacetype"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    Font Face Type?
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    Font Face Type changes the way the text in your forum looks. You may want to change this option to match that of the rest of your web site. Some standards are:
    <ul>
    <li>Arial (nice, clean, legible font)</li>
    <li>Courier (a typewriter font)</li>
    <li>Helvetica (another clean, legible font)</li>
    <li>Sans Serif (Arial & Helvetica are variants of Sans Serif)</li>
    <li>Times New Roman (a book-type font)</li>
    </ul>
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="fontsize"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    What does Font Size mean?
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    <ul>
    <li>1 = 8 point font <b>X-Small</b> (default footer size)</li>
    <li>2 = 10 point font <b>Small</b> (default font size)</li>
    <li>3 = 12 point font <b>Normal</b></li>
    <li>4 = 14 point font <b>Large</b> (default header size)</li>
    <li>5 = 18 point font <b>X-Large</b></li>
    <li>6 = 24 point font <b>XX-Large</b></li>
    <li>7 = 36 point font <b>XXX-Large</b></li>
    </ul>
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="fontdecorations"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>
    What are Font Decorations?
    </b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strFooterFontSize %>">
    <ul>
    <li>无</li>
    <li>闪烁</li>
    <li>删除线</li>
    <li>顶线</li>
    <li>底线</li>
    </ul>
    <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
    </font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="colors"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>What colors may I use?</b></font></td>
  </tr>
  <tr>
      <td bgcolor="<% =strForumCellColor %>">
      <p><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strForumFontColor %>">
      There are a lot of different colors you can choose from, all of which are listed below:</p>
      <blockquote><pre><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
      <font color=aliceblue>aliceblue</font>
      <font color=antiquewhite>antiquewhite</font>
      <font color=aqua>aqua</font>
      <font color=aquamarine>aquamarine</font>
      <font color=azure>azure</font>
      <font color=beige>beige</font>
      <font color=bisque>bisque</font>
      <font color=black>black</font>
      <font color=blanchedalmond>blanchedalmond</font>
      <font color=blue>blue</font>
      <font color=blueviolet>blueviolet</font>
      <font color=brown>brown</font>
      <font color=burlywood>burlywood</font>
      <font color=cadetblue>cadetblue</font>
      <font color=chartreuse>chartreuse</font>
      <font color=chocolate>chocolate</font>
      <font color=coral>coral</font>
      <font color=cornflowerblue>cornflowerblue</font>
      <font color=cornsilk>cornsilk</font>
      <font color=cyan>cyan</font>
      <font color=darkblue>darkblue</font>
      <font color=darkcyan>darkcyan</font>
      <font color=darkgoldenrod>darkgoldenrod</font>
      <font color=darkgray>darkgray</font>
      <font color=darkgreen>darkgreen</font>
      <font color=darkkhaki>darkkhaki</font>
      <font color=darkmagenta>darkmagenta</font>
      <font color=darkolivegreen>darkolivegreen</font>
      <font color=darkorange>darkorange</font>
      <font color=darkorchid>darkorchid</font>
      <font color=darkred>darkred</font>
      <font color=darksalmon>darksalmon</font>
      <font color=darkseagreen>darkseagreen</font>
      <font color=darkslateblue>darkslateblue</font>
      <font color=darkslategray>darkslategray</font>
      <font color=darkturquoise>darkturquoise</font>
      <font color=darkviolet>darkviolet</font>
      <font color=deeppink>deeppink</font>
      <font color=deepskyblue>deepskyblue</font>
      <font color=dimgray>dimgray</font>
      <font color=dodgerblue>dodgerblue</font>
      <font color=firebrick>firebrick</font>
      <font color=floralwhite>floralwhite</font>
      <font color=forestgreen>forestgreen</font>
      <font color=gainsboro>gainsboro</font>
      <font color=ghostwhite>ghostwhite</font>
      <font color=gold>gold</font>
      <font color=goldenrod>goldenrod</font>
      <font color=gray>gray</font>
      <font color=green>green</font>
      <font color=greenyellow>greenyellow</font>
      <font color=honeydew>honeydew</font>
      <font color=hotpink>hotpink</font>
      <font color=indianred>indianred</font>
      <font color=ivory>ivory</font>
      <font color=khaki>khaki</font>
      <font color=lavender>lavender</font>
      <font color=lavenderblush>lavenderblush</font>
      <font color=lawngreen>lawngreen</font>
      <font color=lemonchiffon>lemonchiffon</font>
      <font color=lightblue>lightblue</font>
      <font color=lightcoral>lightcoral</font>
      <font color=lightcyan>lightcyan</font>
      <font color=lightgoldenrod>lightgoldenrod</font>
      <font color=lightgoldenrodyellow>lightgoldenrodyellow</font>
      <font color=lightgray>lightgray</font>
      <font color=lightgreen>lightgreen</font>
      <font color=lightpink>lightpink</font>
      <font color=lightsalmon>lightsalmon</font>
      <font color=lightseagreen>lightseagreen</font>
      <font color=lightskyblue>lightskyblue</font>
      <font color=lightslateblue>lightslateblue</font>
      <font color=lightslategray>lightslategray</font>
      <font color=lightsteelblue>lightsteelblue</font>
      <font color=lightyellow>lightyellow</font>
      <font color=limegreen>limegreen</font>
      <font color=linen>linen</font>
      <font color=magenta>magenta</font>
      <font color=maroon>maroon</font>
      <font color=mediumaquamarine>mediumaquamarine</font>
      <font color=mediumblue>mediumblue</font>
      <font color=mediumorchid>mediumorchid</font>
      <font color=mediumpurple>mediumpurple</font>
      <font color=mediumseagreen>mediumseagreen</font>
      <font color=mediumslateblue>mediumslateblue</font>
      <font color=mediumspringgreen>mediumspringgreen</font>
      <font color=mediumturquoise>mediumturquoise</font>
      <font color=mediumvioletred>mediumvioletred</font>
      <font color=midnightblue>midnightblue</font>
      <font color=mintcream>mintcream</font>
      <font color=mistyrose>mistyrose</font>
      <font color=moccasin>moccasin</font>
      <font color=navajowhite>navajowhite</font>
      <font color=navy>navy</font>
      <font color=navyblue>navyblue</font>
      <font color=oldlace>oldlace</font>
      <font color=olivedrab>olivedrab</font>
      <font color=orange>orange</font>
      <font color=orangered>orangered</font>
      <font color=orchid>orchid</font>
      <font color=palegoldenrod>palegoldenrod</font>
      <font color=palegreen>palegreen</font>
      <font color=paleturquoise>paleturquoise</font>
      <font color=palevioletred>palevioletred</font>
      <font color=papayawhip>papayawhip</font>
      <font color=peachpuff>peachpuff</font>
      <font color=peru>peru</font>
      <font color=pink>pink</font>
      <font color=plum>plum</font>
      <font color=powderblue>powderblue</font>
      <font color=purple>purple</font>
      <font color=red>red</font>
      <font color=rosybrown>rosybrown</font>
      <font color=royalblue>royalblue</font>
      <font color=saddlebrown>saddlebrown</font>
      <font color=salmon>salmon</font>
      <font color=sandybrown>sandybrown</font>
      <font color=seagreen>seagreen</font>
      <font color=seashell>seashell</font>
      <font color=sienna>sienna</font>
      <font color=skyblue>skyblue</font>
      <font color=slateblue>slateblue</font>
      <font color=slategray>slategray</font>
      <font color=snow>snow</font>
      <font color=springgreen>springgreen</font>
      <font color=steelblue>steelblue</font>
      <font color=tan>tan</font>
      <font color=thistle>thistle</font>
      <font color=tomato>tomato</font>
      <font color=turquoise>turquoise</font>
      <font color=violet>violet</font>
      <font color=violetred>violetred</font>
      <font color=wheat>wheat</font>
      <font color=white>white</font>
      <font color=whitesmoke>whitesmoke</font>
      <font color=yellow>yellow</font>
      <font color=yellowgreen>yellowgreen</font>
      </font></pre></blockquote>
      <a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" border="0" align="right"></a>
      </font>
      </td>
  </tr>
</table>
    </td>
  </tr>
</table>

<!--#INCLUDE FILE="inc_footer_short.asp" -->
