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
<!--#include file="config.asp"-->
<!--#include file="inc_functions.asp"-->
<!-- #include file="include/clsUpload.inc" -->

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
   Dim UploadDir
   

   strResult = "Finished..."

   If Request.TotalBytes > 0 Then
     Set objUpload = new clsUpload
     objUpload.ParseRequest
     If (objUpload.Files.Count = 1) Then
        strFolder = objUpload.Form.Item("Folder")
		TopicID = objUpload.Form.Item("TOPIC_ID")
		ReplyID = objUpload.Form.Item("REPLY_ID")
		UploadDir = objUpload.Form.Item("uploaddir")
		if ReplyId = "" then 
		ReplyID = -1
		end if
		MemberID = objUpload.Form.Item("MEMBER_ID")
		faname = objUpload.Files.Item(0).FileName

		boolSuccess = saveDoc()
       	if boolSuccess then 
			strResult = "文件 " & objUpload.Files.Item(0).FileName & " 已上传到服务器"
		end if
     Else
       strResult = "没有文件怎么上传啊？？"
     End If
   End If



   '---------------------------------------------------------------------------
   '- Function    : saveDoc
   '---------------------------------------------------------------------------
   private function saveDoc()
   Dim strRoot
   Dim strFile
   Dim fso
     strRoot = strCookieURL
	 if UploadDir <> "" then
	 	strRoot = Server.MapPath(strCookieURL) & "\" & UploadDir
	else
	 strRoot = Server.MapPath(strRoot) & "\" & strFolder
	 end if

     strFile = strRoot & "\"  & objUpload.Files.Item(0).FileName
     Set fso = Server.CreateObject("Scripting.FileSystemObject")
	 If not fso.FolderExists(strRoot) then

	 	fso.createFolder(strRoot)
	 end if
       if (fso.FileExists(strFile)) then
	   	fso.deleteFile(strFile)
	   end if
	   objUpload.Files.Item(0).Save strRoot
       saveDoc = (fso.FileExists(strFile))
   End Function
%>

<table width="100%" height="100%">
  <tr>
    <td align=center valign=center>
    <div align=center><center>
    <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">

<p>	<%=strResult%></p><br><br>

    <p><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript:history.go(-1)">返回上一页</A></font></p>
    </font>

    </font>
	
    </center></div>
    </td>
  </tr>
</table>
