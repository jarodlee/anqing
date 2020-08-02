<!--#include file="inc_pollmentor.asp"-->
<!--#include file="config.asp"-->
<!--#include file="inc_top_short.asp"-->
<% 	
function remoteIP()	
	
	remoteIP = Request.ServerVariables("HTTP_X_FORWARDED_FOR")	
	if remoteIP = "" then		
		remoteIP = Request.ServerVariables("REMOTE_ADDR")	
	end if
end function

	Dim aTitle
	Dim aQuestion
	Dim aAnswers(8)
	Dim aCount(8)
	'Get active one...
	PollMentor_GetPollInfo -1, aTitle, aQuestion, aAnswers, aCount 
	%>

<TABLE  BORDER=0 bgcolor="<% =strTableBorderColor %>" cellspacing="1" cellpadding="3" width="95%">
    <TR><form method="POST" action="pollmentorres.asp" name="pollForm">
      <TD bgcolor="<% =strCategoryCellColor %>" ALIGN="Left"><font face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>" size="<% =strDefaultFontSize %>"><%=aQuestion%><b>&nbsp;?</b></font></TD>
    </TR>
    <TR>
      <TD BGCOLOR="<% =strForumCellColor %>" ALIGN="LEFT">
<%  
Dim nCount
    Dim nDef
    Dim nTotal
    nTotal = 0
    For nCount=1 To 8 
    	If aAnswers(nCount) <> "" Then
    		nTotal = nTotal + aCount(nCount) %>
        <font face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>" size="<% =strFooterFontSize %>"><input type="radio" value="<%=nCount%>" <%If nDef = False Then %>checked<%End If%> name="R1"><%=aAnswers(nCount)%></font><br>
<%    	nDef = True
    	End If
    Next %>
	</TD>
	</TR><TR>
      <TD BGCOLOR="<% =strForumCellColor %>" ALIGN="CENTER"><font face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>" size="<% =strFooterFontSize %>"><input type="submit" value="投票！" name="B1"><br>
      <a href="pollmentorres.asp">目前结果</a><br>
      <a href="pollmentorhist.asp">清除投票结果</a></font>
	  </TD>
    </TR>
	     </form>
	<tr><td BGCOLOR="<% =strForumCellColor %>">
      <% If nTotal < 2 then %>
	  <font face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>" size="<% =strFooterFontSize %>">到目前共有 <%=nTotal%> 人参加本次投票</font>
	  <% ElseIf nTotal = 0  then%>
	  <font face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>" size="<% =strFooterFontSize %>">到目前为止并没有人参加本次投票</font>
	  <% Else  %>
	  <font face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>" size="<% =strFooterFontSize %>">到目前共有 <%=nTotal%> 人参加本次投票</font>
	  <% End If %>

	</td></tr>	 
</TABLE>
<!--#include file="inc_footer_short.asp"-->