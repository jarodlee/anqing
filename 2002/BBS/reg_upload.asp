<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="inc/Forum_css.asp"-->
</head>
<body bgcolor=<%=Forum_body(4)%> text="#000000" leftmargin="0" topmargin="0">
  <table border="0"  cellspacing="0" cellpadding="0" width=100%>
    <tr>
      <td class=tablebody1>
<form name="form" method="post" action="upfile.asp" enctype="multipart/form-data" >
<input type="hidden" name="filepath" value="uploadFace">
<input type="hidden" name="act" value="upload">
<input type="file" name="file1">
<input type="hidden" name="fname">
<input type="submit" name="Submit" value="上传" onclick="fname.value=file1.value,parent.document.forms[0].Submit.disabled=true,
parent.document.forms[0].Submit2.disabled=true;">
</form>
</td></tr>
</table>
</body>
</html>
<%
conn.close
set conn=nothing
%>