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
            <td bgcolor="<% =strForumCellColor %>" align=center><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>��ʱû���κα���</b><br><a href="<%=request.servervariables("script_name")%>">������̳</a></font></td>
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
