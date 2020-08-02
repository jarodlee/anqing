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
    <td bgcolor="<% =strCategoryCellColor %>"><a name="smilies"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>Smilies</b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>">
    <p><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strForumFontColor %>">
    You've probably seen others use smilies before in email messages or other bulletin 
    board posts. Smilies are keyboard characters used to convey an emotion, such as a smile 
    <img border="0" src="<%=strImageURL %>icon_smile.gif"> or a frown 
    <img border="0" src="<%=strImageURL %>icon_smile_sad.gif">. This bulletin board 
    automatically converts certain text to a graphical representation when it is 
    inserted between brackets [].&nbsp; Here are the smilies that are currently 
    supported by <% =strForumTitle %>:<br>
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
        strsql = "select smile_name from "&strTablePrefix&"smiles where cat_id="&request.querystring("cat_id")
        
        set drs = my_conn.execute(strsql)
        			totalrecordcount = 0
                	do while not drs.eof
                    totalrecordcount = totalrecordcount + 1
                    drs.movenext
                    loop
        set drs =  nothing
        
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
            <td bgcolor="<% =strForumCellColor %>" align=center><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>暂时没有任何表情</b><br><a href="<%=request.servervariables("script_name")%>">返回论坛</a></font></td>
          </tr>

<%		  Response.end
			end if
          set drs=nothing
         %>
          
          </table>
        </td>
        <td>
        <table border="0" align="center">
        <%
        strTempCount = 0
        strsql = "Select smile_name, smile_url, smile_code from "&strTablePrefix&"smiles where cat_id="&request.querystring("cat_id")
        Set drs = Server.CreateObject("ADODB.Recordset")
        drs.open strsql, my_conn, 3
        drs.PageSize = strCount2
        drs.AbsolutePage = 2
        	Do until strTempCount = strCount2
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

<!--#INCLUDE FILE="inc_footer_short.asp" -->
