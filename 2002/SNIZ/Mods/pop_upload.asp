<!--#include file="../config.asp"-->
<head>
<title><% =strForumTitle %></title>
<meta name="copyright" content="This code is Copyright (C) 2000  Michael Anderson">
<Style><!--
a:link    {color:<% =strLinkColor %>;text-decoration:<% =strLinkTextDecoration %>}
a:visited {color:<% =strVisitedLinkColor %>;text-decoration:<% =strVisitedTextDecoration %>}
a:hover   {color:<% =strHoverFontColor %>;text-decoration:<% =strHoverTextDecoration %>}
--></style>
</head>

<BODY background="../<%= strPageBGImage %>" bgColor="<% =strPageBGColor %>" text="<% =strDefaultFontColor %>" link="<% =strLinkColor %>" aLink=<% =strActiveLinkColor %> vLink="<% =strActiveLinkColor %>" onLoad="window.focus()">
<% strCookieURL = Left(Request.ServerVariables("Path_Info"), InstrRev(Request.ServerVariables("Path_Info"), "/Mods/"))
	strDBNTUserName = Request.Cookies(strUniqueID & "User")("Name") %>
<br><br><br>
<div align="center"> <FORM enctype="multipart/form-data" action="UploadEngine.asp" method=POST name="fileup">
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
   <TABLE width="80%" border=0 cellspacing=2 cellpadding=2>
    <TR>

     <TD>
	 Select file to upload...
      <INPUT name="Folder" type="hidden" value="<%= strDBNTUsername %>" >
      <input name="REPLY_ID" type="hidden" value="<% =Request("REPLY_ID") %>">
	  <input name="TOPIC_ID" type="hidden" value="<% =Request("TOPIC_ID") %>">
	  <input name="MEMBER_ID" type="hidden" value="<% =Request("MEMBER_ID") %>">
    </TD>
    </TR>
    <TR>

     <TD>
      <INPUT name="file1" type="FILE" size=20 value="">
     </TD>
	 <td></td>
    </TR>
	<TR>
     <TD ><br>
      <INPUT name="Upload" type=SUBMIT value="Upload.." >
     </TD>
	 <td></td>
    </TR>
  </FORM>
   </TABLE>
   <br><br><a href="Javascript:window.close()">πÿ±’±æ ”¥∞</a>
   </font>
</div>

 </BODY>
</HTML>


