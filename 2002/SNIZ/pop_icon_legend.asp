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
<!--#INCLUDE FILE="config.asp" -->  
<!--#INCLUDE FILE="inc_top_short.asp" -->

<script language="Javascript">
<!-- hide

function insertsmilie(smilieface){

	window.opener.document.PostTopic.Message.value+=smilieface;
}
// -->
</script>


<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td bgcolor="<% =strTableBorderColor %>">
<table border="0" width="100%" cellspacing="1" cellpadding="4">
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="smilies"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>表情符号</b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>">
    <p><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strForumFontColor %>">
    你或许在电子邮件或其他论坛看过别人使用表情符号，这是指一些代表情绪的键盘符号， 像是一个笑脸<img border="0" src="<%=strImageURL %>icon_smile.gif">或是难过的脸<img border="0" src="<%=strImageURL %>icon_smile_sad.gif">。这个论坛系统可以自动将这些符号转换成图形，譬如说文章中的这个符号 [:D] 会自动转换成这个图形<img border="0" src="<%=strImageURL %>icon_smile_big.gif">所有的表情图案都要记得加上[]中括号，下面就是<% =strForumTitle %>所有提供的表情图案，如果你有更丰富的表情图案建议，请主动提供：<br>
    <table border="0" align="center" cellpadding="5">
      <tr valign="top">
        <td>
        <table>
        <%
        if request.querystring("cat_id") = "" then
        %>
        <table bgcolor="<%=strTableBorderColor%>" border=1 cellpadding=1 cellspacing=1>
        <%
        	strsql = "select cat_id, cat_name from "&strTablePrefix&"smile_cat"
            set catRS1 = my_conn.execute(strsql)
            Do until catRS1.eof
        %>
        <tr bgcolor="<%=strForumCellColor%>">
        	<td><font face="<%=strDefaultFontFace%>" size=<%=strDefaultFontSize+1%>>
            <a href="<%=request.servervariables("script_name")%>?cat_id=<%=catRS1("cat_id")%>"><%=catRS1("cat_name")%></a>
            </td>
        </tr>
        <%
        	catRS1.movenext
            loop
            set carRS1=nothing
        %>
        </table>
        <%
        else
        %>
        <table border="0" align="center" cellpadding="5">
        <tr valign="top">
          <td>
        <table border="0" align="center">
        <%
        strsql = "select count(ID) AS TotalCount from " & strTablePrefix & "smiles "
		'strsql = strsql & "WHERE " & strTablePrefix & "SMILES.CAT_ID>" & 0
		strsql = strsql & "WHERE " & strTablePrefix & "SMILES.CAT_ID=" & request.querystring("cat_id")
	set drsCount = server.CreateObject("ADODB.RecordSet")
	drsCount.CacheSize= 30

    drsCount.Open strsql, my_Conn,3
	
        			totalrecordcount = 0
                    totalrecordcount = drsCount("TotalCount")
        set drsCount =  nothing
        
        If totalrecordcount mod 2 = 0 then   
        strCount1 = totalrecordcount / 2   
        strCount2 = strCount1
        else   
        strCount1 = (totalrecordcount / 2) + 0.5   
        strCount2 = strCount1 - 1
        end if
        strTempCount = 0
        
        strsql = "Select smile_name, smile_url, smile_code from "&strTablePrefix&"smiles where cat_id="&request.querystring("cat_id")
        Set drs = Server.CreateObject("ADODB.Recordset")
        drs.open strsql, my_conn, 3
      	if not (drs.eof or drs.bof) then
			drs.PageSize = strCount1
        	drs.AbsolutePage = 1
        	Do until strTempCount = strCount1
            smile_name = drs("smile_name")
            smile_url = drs("smile_url")
            smile_code = drs("smile_code")
           %>
          <tr>
            <td bgcolor="<% =strForumCellColor %>"><a href="Javascript:insertsmilie('[<%=smile_code%>]');"><img border="0" hspace="10" src="<%=strImageURL %><%=smile_url%>"></a></td>
            <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><%=smile_name%></font></td>
            <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">[<%=smile_code%>]</font></td>
          </tr>
          <%
          strTempCount = strTempCount + 1
          drs.movenext
          loop
	else%>
          <tr>
            <td bgcolor="<% =strForumCellColor %>" align=center><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>目前没有任何图案</b><br><a href="<%=request.servervariables("script_name")%>">返回</a></font></td>
          </tr>

<%		 
			Response.end
		end if
          
         %>
          
          </table>
        </td>
        <td>
        <table border="0" align="center">
        <%
        strsql = "Select smile_name, smile_url, smile_code from "&strTablePrefix&"smiles where cat_id="&request.querystring("cat_id")
        Set drs = Server.CreateObject("ADODB.Recordset")
        drs.open strsql, my_conn, 3
			drs.PageSize = strCount1
			if strcount2 > 0 then
			drs.AbsolutePage = 2
			strTempCount = 0
        	Do until (strTempCount = strCount2) or (drs.eof)
            smile_name = drs("smile_name")
            smile_url = drs("smile_url")
            smile_code = drs("smile_code")
        %>
        	<tr>
            <td bgcolor="<% =strForumCellColor %>"><a href="Javascript:insertsmilie('[<%=smile_code%>]');"><img border="0" hspace="10" src="<%=strImageURL %><%=smile_url%>"></a></td>
            <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><%=smile_name%></font></td>
            <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">[<%=smile_code%>]</font></td>
          </tr>
          <%
          strTempCount = strTempCount + 1
          drs.movenext
          loop
		  end if
          set drs = nothing
          %>
          </table>
          <%end if%>
          </td>
      </tr>
    </table></p>
    </td>
  </tr>
</table>
    </td>
  </tr>
</table>
    </td>
  </tr>
</table>
<p><div align="center"><a href="javascript:history.go(-1)">返回</a></div></p>
<!--#INCLUDE FILE="inc_footer_short.asp" -->
