<!--#include file="../config.asp"-->
<!--#include file="../inc_functions.asp"-->
<!-- #include file="../Include/clsUpload.inc" -->

<%
   '--- Object to parse the input ---
   Dim objUpload

   Dim strFolder
   Dim MaxFileSize
   Dim boolAllowedType
   Dim strAllowedExtensions
   Dim strResult
   Dim faname 
   Dim boolSuccess
   
	set my_Conn= Server.CreateObject("ADODB.Connection")
	my_Conn.Open strConnString
	
	Set objDict = Server.CreateObject("Scripting.Dictionary")
	set objRec = my_Conn.execute("SELECT * FROM " & strTablePrefix & "MODS WHERE M_NAME = 'Attachment' OR M_CODE = 'Attachment'")

	while not objRec.EOF 
	objDict.Add objRec.Fields.Item("m_code").Value, objRec.Fields.Item("m_value").Value
	objRec.moveNext
	wend 
	
	MaxFileSize = cLng(objDict.Item("faMaxSize"))*512
	strAllowedExtensions = objDict.Item("faExtensions")
	boolAllowUploads = cint(objDict.Item("Attachment"))

	objRec.Close
	set objDict = nothing
	set my_Conn = nothing


   strResult = "Finished..."

   If Request.TotalBytes > 0 Then
     Set objUpload = new clsUpload
     objUpload.ParseRequest
     If (objUpload.Files.Count = 1) Then
        strFolder = objUpload.Form.Item("Folder")
		TopicID = objUpload.Form.Item("TOPIC_ID")
		ReplyID = objUpload.Form.Item("REPLY_ID")
		if ReplyId = "" then 
		ReplyID = -1
		end if
		MemberID = objUpload.Form.Item("MEMBER_ID")
		faname = objUpload.Files.Item(0).FileName
        boolAllowedType = Instr(strAllowedExtensions, Right(objUpload.Files.Item(0).FileName,4)) > 0
	    intFileLen = len(objUpload.Files.Item(0).Blob)
		if intFileLen > MaxFileSize then
			strResult = "File is too Large to upload"
		else
	        if boolAllowedType then
				boolSuccess = saveDoc()
				If boolSuccess = "True" Then
	          	intoDB()
  
	          	strResult = "The File " & objUpload.Files.Item(0).FileName & " was uploaded to the Server"
	        	End If
			else
				strResult = "File Type Not Allowed"
			end if
		end if
     Else
       strResult = "No file to upload."
     End If
   End If


   '---------------------------------------------------------------------------
   '- Function    : intoDB
   '- Description : inserts the document information in the database
   '---------------------------------------------------------------------------
   Private Function intoDB()
   	Dim strSQL
   	set my_Conn= Server.CreateObject("ADODB.Connection")
	my_Conn.Open strConnString

     strSQL = "INSERT INTO FORUM_USERFILES (MEMBER_ID,F_FILENAME,F_FILESIZE,F_TOPIC_ID,F_REPLY_ID) VALUES("
     strSQL = strSQL  & MemberID & ", "
     strSQL = strSQL & "'" & objUpload.Files.Item(0).FileName & "', 0, "
     strSQL = strSQL  & TopicID & ", "
     strSQL = strSQL & ReplyID & ")"
     my_Conn.execute( strSQL )
	 'response.write strSQL
   End Function

   '---------------------------------------------------------------------------
   '- Function    : saveDoc
   '---------------------------------------------------------------------------
   private function saveDoc()
   Dim strRoot
   Dim strFile
   Dim fso
     strRoot = replace(strCookieURL,"/mods/","/",1,-1,1) & "uploaded"
	 strRoot = Server.MapPath(strRoot) & "\" & strFolder

     strFile = strRoot & "\"  & objUpload.Files.Item(0).FileName
     Set fso = Server.CreateObject("Scripting.FileSystemObject")
	 If not fso.FolderExists(strRoot) then

	 	fso.createFolder(strRoot)
	 end if
       objUpload.Files.Item(0).Save strRoot
       saveDoc = (fso.FileExists(strFile))
   End Function
%>
<html>

<head>
<title><% =strForumTitle %></title>
<meta name="copyright" content="This code is Copyright (C) 2000  Michael Anderson">
<Style><!--
a:link    {color:<% =strLinkColor %>;text-decoration:<% =strLinkTextDecoration %>}
a:visited {color:<% =strVisitedLinkColor %>;text-decoration:<% =strVisitedTextDecoration %>}
a:hover   {color:<% =strHoverFontColor %>;text-decoration:<% =strHoverTextDecoration %>}
--></style>
</head>

<BODY bgColor="<% =strPageBGColor %>" text="<% =strDefaultFontColor %>" link="<% =strLinkColor %>" aLink=<% =strActiveLinkColor %> vLink="<% =strActiveLinkColor %>" onLoad="window.focus()">

<table width="100%" height="100%">
  <tr>
    <td align=center valign=center>
    <div align=center><center>
    <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">

<p>	<%=strResult%></p><br><br>

    <p><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript:onClick= addfile()">πÿ±’±æ ”¥∞</A></font></p>
    </font>
    </center></div>
    </td>
  </tr>
</table>
<script language="JavaScript">
<!--
function addfile()
{
	var fname, resultString
	resultString ='<%= boolSuccess %>';
	fname = '<%= faname %>';
	if (resultString)
	window.opener.document.PostTopic.Message.value+= '[url="pop_download.asp?mode=Edit&dir='  + '<%=strFolder%>' + '&file=' + fname + '"]download[/url]';
	window.close();
}

//-->
</script>

</BODY>
</HTML>
