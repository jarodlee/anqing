<%
'########## Snitz Forums 2000 Version 3.1 SR4 ####################
'#                                                               #
'#  �����޸�: ��Դ����վ                                         #
'#  �����ʼ�: cgier@21cn.com                                     #
'#  ��ҳ��ַ: http://www.sdsea.com                               #
'#            http://www.99ss.net                                #
'#            http://www.cdown.net                               #
'#	     http://www.wzdown.com                               #
'#	     http://www.13888.net                                #
'#  ��̳��ַ:http://ubb.yesky.net                                #
'#  ����޸�����: 2001/03/12    ���İ汾��Version 3.1 SR4        #
'#################################################################
'# ԭʼ��Դ                                                      #
'# Snitz Forums 2000 Version 3.1 SR4                             #
'# Copyright 2000 http://forum.snitz.com - All Rights Reserved   #
'#################################################################
'#����Ȩ������                                                   #
'#                                                               #
'# ������Ϊ��������(shareware)�ṩ������վ���ʹ�ã�����Ƿ��޸�,#
'# ת�أ�ɢ��������������ͼ����Ϊ��������ɾ����Ȩ������          #
'# ���������վ��ʽ����������ű�������֪ͨ���ǣ��Ա������ܹ�֪��#
'# ������ܣ�����������վ�������ǵ����ӣ�ϣ���ܸ��������лл��  #
'#################################################################
'# �����������ǵ��Ͷ��Ͱ�Ȩ����Ҫɾ�����ϵİ�Ȩ�������֣�лл����#
'# �����κ������뵽���ǵ���̳��������                            #
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
			strResult = "�ļ� " & objUpload.Files.Item(0).FileName & " ���ϴ���������"
		end if
     Else
       strResult = "û���ļ���ô�ϴ�������"
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

    <p><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="JavaScript:history.go(-1)">������һҳ</A></font></p>
    </font>

    </font>
	
    </center></div>
    </td>
  </tr>
</table>
