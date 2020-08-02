<% Response.Buffer = True %>
<!--#include file="inc_pollmentor.asp"-->
<!--#include file="config.asp"-->
<!--#include file="inc_top_short.asp"-->
<%

Dim sTitle
Dim vQuestion
Dim vAnswers(8)
Dim vCount(8)
Dim vPercent(8)
	
scriptname = Request.ServerVariables("HTTP_REFERER")
'Some before doings...
'Lets get the current poll...

sID = Request.QueryString("id")
If sID = "" Then
	sID = -1
End If

Dim sError

'Are we trying to vote?

If Request.Form("R1") <> "" Then
	'First try to vote...
	sError = PollMentor_TryToVote(sID,Request.Form("R1") )
End If

'Get active one...
Dim nRet
nRet = PollMentor_GetPollInfo( sID, sTitle, sQuestion, vAnswers, vCount )
%>

<div align="center"><font face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>" size="<% =strDefaultFontSize %>"><%=sError%></FONT><br>



 <TABLE BORDER=0 cellspacing="1" cellpadding="3" bgcolor="<% =strTableBorderColor %>" >
    <TR >
      <TD bgcolor="<% =strCategoryCellColor %>" COLSPAN=3 align="center"><b><font face="<% =strDefaultFontFace %>" color="<% =strCategoryFontColor %>" size="<% =strDefaultFontSize %>"><%=sQuestion%>&nbsp;<b>?</b></font></b></TD>
    </TR><%

'First of all get max value and nLowValue
nMaxValue = 0
For nCount=1 To 8 
	If vAnswers(nCount)<>"" And vCount(nCount)>nMaxValue Then
		nMaxValue = vCount(nCount)
	End If
Next	

If nMaxValue = 0 Then	
	nMaxValue = 1
End If


nMaxWidth = 60 'This is number of pixels for maxvalue

nTotal = 0

'1. Go through all and get total
For nCount=1 To 8 
	If vAnswers(nCount)<>"" Then
		nTotal = nTotal + vCount(nCount)
	End If
Next

If nTotal = 0 Then
	nTotal = 1
End If

'2. Go through all and get percent 
For nCount=1 To 8 
	If vAnswers(nCount)<>"" Then
		vPercent(nCount) = FormatNumber(vCount(nCount)/nTotal*100,1)
	End If
Next

For nCount=1 To 8 
If vAnswers(nCount)<>"" Then
	nThisVal = FormatNumber(vCount(nCount)/nMaxValue * nMaxWidth,0)
%>
    <TR>
      <TD bgcolor="<% =strForumCellColor %>" COLSPAN=2 Align= LEFT><font face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>" size="<% =strFooterFontSize %>"><%=vAnswers(nCount)%></FONT></TD>
     
    </TR>
    <TR>
      <TD bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>" size="<% =strFooterFontSize %>"><img src="<%=strImageURL %>bar.gif" width="<%=nThisVal%>" height="10"> (<%=vPercent(nCount)%> %)</font></TD>
      <TD bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>" size="<% =strFooterFontSize %>"><%=vCount(nCount)%> 张选票</font></TD>
      </TR><%
	End If
Next
%>
    <TR>
      <TD align="left" bgcolor="<% =strForumCellColor %>" COLSPAN=3>
	  <font face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>" size="<% =strFooterFontSize %>">总选票数：<%=nTotal%></font>
	  </TD>
    </TR>
    <TR>
      <TD bgcolor="<% =strForumCellColor %>" COLSPAN=3><div align="center"><font face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>" size="<% =strFooterFontSize %>">
            
			<p><a href="pollmentorhist.asp" target="_self">之前投票结果</a></p>
			
						  <a href="javascript:history.go(-1)">回上页</a>
			</font></div>
		</TD>
		      </TR>
</TABLE>

<!--#include file="inc_footer_short.asp"-->