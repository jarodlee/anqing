<!--#INCLUDE FILE="../config.asp" -->
<!--#INCLUDE FILE="../inc_functions.asp" -->
<html>

<head>
<title>ReadMe Viewer - by - AnimeDJ.com</title>
<meta name="copyright" content="This code is Copyright (C) 2000  Michael Anderson">
<script language="JavaScript">
<!-- hide from JavaScript-challenged browsers
function openWindow(url) {
  popupWin = window.open(url,'new_page','width=400,height=400')
}
function openWindow2(url) {
  popupWin = window.open(url,'new_page','width=400,height=450')
}
function openWindow3(url) {
  popupWin = window.open(url,'new_page','width=400,height=450,scrollbars=yes')
}
function openWindow4(url) {
  popupWin = window.open(url,'new_page','width=400,height=525')
}
// done hiding -->
</script>
<style type=text/css><!--
a:link    {color:<% =strLinkColor %>;text-decoration:<% =strLinkTextDecoration %>}
a:visited {color:<% =strVisitedLinkColor %>;text-decoration:<% =strVisitedTextDecoration %>}
a:hover   {color:<% =strHoverFontColor %>;text-decoration:<% =strHoverTextDecoration %>}
input.radio {background: <% = strPopUpTableColor %>; color:#000000}
--></style>
</HEAD>

<body bgColor="<% =strPageBGColor %>" text="<% =strDefaultFontColor %>" link="<% =strLinkColor %>" aLink="<% =strActiveLinkColor %>" vLink="<% =strActiveLinkColor %>">



<%
    Dim rfile1, rfile2, thisline, rfile
    
    rfile = request.querystring("readme") & ".txt"
%>

<body bgcolor="#FFFFFF">
<table width="90%" border="0" cellspacing="1" cellpadding="4" bgcolor="#000000" align="center">
  <tr bgcolor="<% =strCategoryCellColor %>" align="center"> 
    <td> <font face="ËÎÌו, Arial, Helvetica, sans-serif" size="3"><b><a href="../admin_home.asp">Go 
      Back</a></b></font></td>
  </tr>
  <tr bgcolor="<% =strForumCellColor %>"> 
    <td>
      <font face="ËÎÌו, Arial, Helvetica, sans-serif" size="2">
<%
    Set rfile1 = CreateObject("Scripting.FileSystemObject")
    Set rfile2 = rfile1.OpenTextFile(server.mappath(rfile), 1, False)

    do while not rfile2.AtEndOfStream
        response.write HTMLEncode(rfile2.readline) & "<br>"
    loop
    
    rfile2.Close
    set rfile2=nothing
    set rfile1=nothing
%>    

      </FONT>    
    </td>
  </tr>
</table>
</body>
</html>
