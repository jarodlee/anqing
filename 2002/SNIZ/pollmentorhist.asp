<%
Response.Buffer = True
%>
<!--#include file="config.asp"-->
<!--#include file="inc_pollmentor.asp"-->
<!--#include file="inc_top_short.asp"-->
<%
	scriptname = Request.ServerVariables("HTTP_REFERER")
	Dim strTrue
	Dim strFalse
	
	If strDBType = "sqlserver" Then
		strTrue = 	"1"
		strFalse = "0"
	Else
		strTrue = "True"
		strFalse = "False"
	End If

Dim oConn, oRS
Set oConn = PollMentor_GetDatabaseConn()
Set oRS = oConn.Execute("select * from " & strTablePrefix & "poll where active=" & strFalse )
%>
	  <div align="center"><p><font face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>" size="<% =strDefaultFontSize %>">点选投票主题来观看结果</font></p>

      <table border="0" bgcolor="<% =strTableBorderColor %>" cellspacing="1" cellpadding="3" >
        <tr>
          <td align="center" bgcolor="<% =strCategoryCellColor %>"><font face="<% =strDefaultFontFace %>" color="<% =strCategoryFontColor %>" size="<% =strDefaultFontSize %>"><b>之前的投票</b></font></td>
        </tr>
<%
Dim nCount, nTotal
While Not oRS.EOF
	nTotal = 0
	For nCount = 1 To 8
		If oRS("answer" & CStr(nCount)) <> "" Then
			nTotal=nTotal + oRS("count" & CStr(nCount) )
		End If
	Next
%>        
        <tr wrap>
          <td align="left"  bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>" size="<% =strFooterFontSize %>"><a href="pollmentorres.asp?id=<%=oRS("id")%>"><%=oRS("question")%></a><br>（选票数：<%=nTotal%>）</font></td>
        </tr>
<%
	oRS.MoveNext
Wend
%>     
      <TD bgcolor="<% =strForumCellColor %>" COLSPAN=3><div align="center"><font face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>" size="<% =strFooterFontSize %>">
            <br><br>
						  <a href="javascript:history.go(-1)">回上页</a>
			</font></div>
		</TD>
   
      </table>
	  </div>

<!--#include file="inc_footer_short.asp"-->

