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
    <td bgcolor="<% =strCategoryCellColor %>"><a name="smilies"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>�������</b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>">
    <p><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strForumFontColor %>">
    ������ڵ����ʼ���������̳��������ʹ�ñ�����ţ�����ָһЩ���������ļ��̷��ţ� ����һ��Ц��<img border="0" src="<%=strImageURL %>icon_smile.gif">�����ѹ�����<img border="0" src="<%=strImageURL %>icon_smile_sad.gif">�������̳ϵͳ�����Զ�����Щ����ת����ͼ�Σ�Ʃ��˵�����е�������� [:D] ���Զ�ת�������ͼ��<img border="0" src="<%=strImageURL %>icon_smile_big.gif">���еı���ͼ����Ҫ�ǵü���[]�����ţ��������<% =strForumTitle %>�����ṩ�ı���ͼ����������и��ḻ�ı���ͼ�����飬�������ṩ��<br>
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
            <td bgcolor="<% =strForumCellColor %>" align=center><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>Ŀǰû���κ�ͼ��</b><br><a href="<%=request.servervariables("script_name")%>">����</a></font></td>
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
<p><div align="center"><a href="javascript:history.go(-1)">����</a></div></p>
<!--#INCLUDE FILE="inc_footer_short.asp" -->
