<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<script>
if (top.location==self.location){
	top.location="index.asp"
}
</script>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="inc/Forum_css.asp"-->
</head>
<%
dim dateupnum
dateupnum=request.cookies("dateupnum")
if dateupnum ="" then dateupnum=0
dateupnum=int(dateupnum)
%>

<body <%=Forum_body(11)%>>
<form name="form" method="post" action="saveannouce_upfile.asp?boardid=<%=request("boardid")%>" enctype="multipart/form-data">
<table width="100%" border=0 cellspacing=0 cellpadding=0>
<tr><td class=tablebody2 valign=top height=30>
<%if Cint(GroupSetting(7))=0 then%>
��û���ڱ���̳�ϴ��ļ���Ȩ��
<%else%>
<input type="hidden" name="filepath" value="UploadFile/">
<input type="hidden" name="act" value="upload">
<input type="hidden" name="fname">
<input type="file" name="file1">
<input type="submit" name="Submit" value="�ϴ�" onclick="fname.value=file1.value,parent.document.forms[0].Submit.disabled=true,
parent.document.forms[0].Submit2.disabled=true;">
<font color=<%=Forum_body(8) %> >���컹���ϴ�<%=GroupSetting(50)-dateupnum%>��</font>��
  ��̳���ƣ�һ��<%=GroupSetting(40)%>����һ��<%=GroupSetting(50)%>��,ÿ��<%=GroupSetting(44)%>K
<%end if%>
</td></tr>
</table>
</form>
</body>
</html>
<%
conn.close
set conn=nothing
%>