<!--#include FILE="conn.asp"-->
<!--#include FILE="inc/const.asp"-->
<!--#include FILE="upload.inc"-->
<html>
<head>
<title>文件上传</title>
<!--#include file=inc/Forum_css.asp-->
</head>
<body  topmargin="0" leftmargin="0">
<table width=100% border=0  cellspacing="0" cellpadding="0"><tr><td class=tablebody1 width=100% height=100% >
<script>
parent.document.forms[0].Submit.disabled=false;
parent.document.forms[0].Submit2.disabled=false;
</script>
<%
dim upload_type
dim upload,file,formName,formPath,iCount,filename,fileExt
dim ranNum
upload_type=0	'上传方法，0＝无组件，1＝Lyfupload，2＝Aspupload3.0，3＝chinaaspupload

if session("upface")="done" then
	response.write "您已经上传了头像"
	response.end
end if

'******************
'删除用户旧的头像
dim filepaths,objFSO,upface
if  founduser then
on error resume next
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
set rs=conn.execute("select face from [user] where userid="&userid)
    upface=rs(0)
if instr(upface,"uploadFace")>0  then
filepaths=Server.MapPath(""&upface&"")
	if objFSO.fileExists(filepaths) then
	'response.write "删除"&filepaths
	objFSO.DeleteFile(filepaths)
	end if
end if
set rs=nothing
end if
'******************
'删除用户旧的头像

select case upload_type
case 0
	call upload_0()
case 1
	call upload_1()
case 2
	call upload_2()
case else
	response.write "本系统未开放插件功能"
	response.end
end select

'===========无组件上传(upload_0)====================
sub upload_0()
set upload=new UpFile_Class ''建立上传对象
upload.GetDate (int(Forum_Setting(56))*1024)   '取得上传数据,不限大小
iCount=0
 
if upload.err > 0 then
    select case upload.err
	case 1
	Response.Write "请先选择你要上传的文件　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
	case 2
	Response.Write "图片大小超过了限制 "&Forum_Setting(56)&"K　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
	end select
	exit sub
	else
formPath=upload.form("filepath")
 ''在目录后加(/)
 if right(formPath,1)<>"/" then formPath=formPath&"/" 

for each formName in upload.file ''列出所有上传了的文件
 set file=upload.file(formName)  ''生成一个文件对象
 if file.filesize<100 then
 	response.write "请先选择你要上传的图片　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
	response.end
 end if

 fileExt=lcase(file.FileExt)
if CheckFileExt(fileEXT)=false then
 	response.write "文件格式不正确　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
	response.end
end if

randomize
ranNum=int(90000*rnd)+10000
 filename=formPath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&fileExt
 if file.FileSize>0 then         ''如果 FileSize > 0 说明有文件数据
  file.SaveToFile Server.mappath(filename)   ''保存文件
'  response.write file.FilePath&file.FileName&" ("&file.FileSize&") => "&formPath&File.FileName&" 成功!<br>"
 response.write "<script>parent.document.forms[0].myface.value='"&FileName&"'</script>"
  iCount=iCount+1
 end if
 set file=nothing
next
set upload=nothing
session("upface")="done"
Htmend iCount&" 个文件上传结束!"

end if
end sub

'===========Lyfupload组件上传(upload_1)=========================
sub upload_1()
	dim obj,filename,fileExt_a
	dim ss
	Set obj = Server.CreateObject("LyfUpload.UploadFile")
	'大小
    	obj.maxsize=int(Forum_Setting(56))*1024
	'类型
    	obj.extname="gif,jpg,bmp,jpeg"
	'重命名
	formPath=obj.request("filepath")
	'在目录后加(/)
	if right(formPath,1)<>"/" then formPath=formPath&"/" 
	if obj.request("fname")="" or isnull(obj.request("fname")) then
	Response.Write "请先选择你要上传的文件　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
	exit sub
	end if
	randomize
	ranNum=int(90000*rnd)+10000
	fileExt_a=split(obj.request("fname"),".")
	fileExt=lcase(fileExt_a(ubound(fileExt_a)))
	filename=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum
	filename=filename&"."&fileExt

	'response.end
	ss=obj.SaveFile("file1",Server.MapPath(formPath), true,filename)

	if ss= "3" then
   		Response.Write ("文件名重复![ <a href=# onclick=history.go(-1)>重新上传</a> ]")
		response.end
	elseif ss= "0" then
   		Response.Write ("文件尺寸过大![ <a href=# onclick=history.go(-1)>重新上传</a> ]")
		response.end
	elseif ss = "1" then
 		Response.Write ("文件不是指定类型文件![ <a href=# onclick=history.go(-1)>重新上传</a> ]")
		response.end
	elseif ss = "" then
 		Response.Write ("文件上传失败![ <a href=# onclick=history.go(-1)>重新上传</a> ]")
		response.end
	else
		Response.Write "图片上传成功!" 
		response.write "<script>parent.document.forms[0].myface.value='" &formPath&filename & "'</script>"
		session("upface")="done"
		response.end
	end if
	set obj=nothing
end sub



''===========================Aspupload3.0组件上传============================
sub upload_2()
dim Count
on error resume next
Set Upload = Server.CreateObject("Persits.Upload") 
Upload.OverwriteFiles = false   '不能复盖
Upload.IgnoreNoPost = True
Upload.SetMaxSize int(Forum_Setting(56))*1024, True	 '限制大小

Count = Upload.Save
If Err.Number = 8 Then 
   Response.Write "文件大小超过了限制 "&Forum_Setting(56)&"K　[ <a href=# onclick=history.go(-1)>重新上传</a> ]" 
Else 
   If Err <> 0 Then 
      Response.Write "错误信息: " & Err.Description 
   Else
		If Count < 1 Then 
		Response.Write "请先选择你要上传的文件　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
		Response.End
		End If   
	formPath=Upload.form("filepath")
	if right(formPath,1)<>"/" then formPath=formPath&"/" 
	For Each file in Upload.Files	'列出所有上传文件
	fileExt=lcase(replace(File.ext,".",""))
	'判断文件类型
	if CheckFileExt(fileEXT)=false then
 	response.write "文件格式不正确　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
	response.end
	end if
	'文件变量付值
	randomize
	ranNum=int(90000*rnd)+10000
	filename=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&File.ext
	filename=formPath&filename
	file.saveas Server.MapPath(filename)	'上传保存文件
	response.Write "图片上传成功!" 
	response.write "<script>parent.document.forms[0].myface.value='" &filename& "'</script>"
	session("upface")="done"
	Next
   End If 
End If
set Upload =nothing

end sub

sub HtmEnd(Msg)
 response.write "图片上传成功!"
 response.end
end sub

'判断文件类型是否合格
Private Function CheckFileExt (fileEXT)
dim Forumupload
Forumupload="gif,jpg,bmp,jpeg"
Forumupload=split(Forumupload,",")
	for i=0 to ubound(Forumupload)
		if lcase(fileEXT)=lcase(trim(Forumupload(i))) then
			CheckFileExt=true
			exit Function
		else
			CheckFileExt=false
		end if
	next
End Function

%>
</td></tr></table>
</body>
</html>