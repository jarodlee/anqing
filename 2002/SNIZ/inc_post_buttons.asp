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
<script language="Javascript">
<!-- hide

function insertsmilie(smilieface){

	PostTopic.Message.value+=smilieface;
}
// -->
</script>
<tr>
<td bgColor="<% =strPopUpTableColor %>" align=right rowspan=2 valign=top>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">��ʽ��</font>
</td>
<td bgColor="<% =strPopUpTableColor %>" align=left>
<a href="Javascript:bold();"><img src="<%=strImageURL %>icon_editor_bold.gif" width="22" height="22" alt="����" border="0"></a><a href="Javascript:italicize();"><img src="<%=strImageURL %>icon_editor_italicize.gif" width="23" height="22" alt="б��" border="0"></a><a href="Javascript:underline();"><img src="<%=strImageURL %>icon_editor_underline.gif" width="23" height="22" alt="�»���" border="0"></a>
<a href="Javascript:center();"><img src="<%=strImageURL %>icon_editor_center.gif" width="22" height="22" alt="����" border="0"></a>
<a href="Javascript:hyperlink();"><img src="<%=strImageURL %>icon_editor_url.gif" width="22" height="22" alt="���볬������" border="0"></a><a href="Javascript:email();"><img src="<%=strImageURL %>icon_editor_email.gif" width="23" height="22" alt="�����ʼ���ַ" border="0"></a><a href="Javascript:image();"><img src="<%=strImageURL %>icon_editor_image.gif" width="23" height="22" alt="����ͼƬ" border="0"></a>
<a href="Javascript:showcode();"><img src="<%=strImageURL %>icon_editor_code.gif" width="22" height="22" alt="�������" border="0"></a><a href="Javascript:quote();"><img src="<%=strImageURL %>icon_editor_quote.gif" width="23" height="22" alt="��������" border="0"></a><a href="Javascript:list();"><img src="<%=strImageURL %>icon_editor_list.gif" width="23" height="22" alt="�������" border="0"></a>
</td>
</tr><tr>
<td bgColor="<% =strPopUpTableColor %>" align=left>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    <select name="font" onChange="showfont(this.options[this.selectedIndex].value)">
	<option value="����" selected>����</option>
	<option value="Arial">Arial</option>
	<option value="Book Antiqua">Book Antiqua</option>
	<option value="Century Gothic">Century Gothic</option>
	<option value="Comic Sans MS">Comic Sans MS</option>
	<option value="Courier New">Courier New</option>
	<option value="Georgia">Georgia</option>
	<option value="Impact">Impact</option>
	<option value="Tahoma">Tahoma</option>
	<option value="Times New Roman">Times New Roman</option>
	<option value="Trebuchet MS">Trebuchet MS</option>
	<option value="Script MT Bold">Script MT Bold</option>
	<option value="Stencil">Stencil</option>
	<option value="Verdana">Verdana</option>
	<option value="Lucida Console">Lucida Console</option>
</select>&nbsp;
<select name="size" onChange="showsize(this.options[this.selectedIndex].value)">
	<option value="1">1</option>
	<option value="2">2</option>
	<option value="3" selected>3</option>
	<option value="4">4</option>
	<option value="5">5</option>
	<option value="6">6</option>	
</select>&nbsp;
<select name="color" onChange="showcolor(this.options[this.selectedIndex].value)">
	<option value="black" selected>��ɫ</option>
	<option value="red">��ɫ</option>
	<option value="yellow">��ɫ</option>
	<option value="pink">�ۺ�ɫ</option>
	<option value="green">��ɫ</option>
	<option value="orange">��ɫ</option>
	<option value="purple">��ɫ</option>
	<option value="blue">��ɫ</option>
	<option value="beige">�׻�ɫ</option>
	<option value="brown">��ɫ</option>
	<option value="teal">����ɫ</option>
	<option value="navy">����ɫ</option>
	<option value="maroon">���Ϻ�ɫ</option>
	<option value="limeGreen">������ɫ</option>
</select></td>
</tr>
<% if lcase(strIcons) = "1" then %>
<tr><td bgColor="<% =strPopUpTableColor %>" align=right rowspan=1 valign=top>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">������ţ�</font><br>
<a href="JavaScript:openWindow5('pop_icon_legend.asp')">����������</a>
</td>
<td bgColor="<% =strPopUpTableColor %>" align=left>
<%
        strsql = "select count(ID) AS TotalCount from " & strTablePrefix & "smiles "
		'strsql = strsql & "WHERE " & strTablePrefix & "SMILES.CAT_ID>" & 0
		strsql = strsql & "WHERE " & strTablePrefix & "SMILES.CAT_ID=" & 1
	set drsCount = server.CreateObject("ADODB.RecordSet")
	drsCount.CacheSize= 10

    drsCount.Open strsql, my_Conn,3
	
        			totalrecordcount = 0
                    totalrecordcount = drsCount("TotalCount")
        set drsCount =  nothing
        strTempCount = 1
        
        strCount1 = totalrecordcount
        strsql = "Select smile_name, smile_url, smile_code from "&strTablePrefix&"smiles where cat_id="& 1
        Set drs = Server.CreateObject("ADODB.Recordset")
        drs.open strsql, my_conn, 3
      	if not (drs.eof or drs.bof) then
			drs.PageSize = strCount1
        	drs.AbsolutePage = 1
        	Do until strTempCount = strCount1+1
            smile_name = drs("smile_name")
            smile_url = drs("smile_url")
            smile_code = drs("smile_code")
          %>
          <%if strTempCount mod 15 = 0 then %>
          <a href="Javascript:insertsmilie('[<%=smile_code%>]');"><img border="0" src="<%=strImageURL %><%=smile_url%>" alt="<%=smile_name%>"></a><br>
          <%else%>
          <a href="Javascript:insertsmilie('[<%=smile_code%>]');"><img border="0" src="<%=strImageURL %><%=smile_url%>" alt="<%=smile_name%>"></a>
          <%end if%>
          <%
          strTempCount = strTempCount + 1
          drs.movenext
          loop
		  else%>
          <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">û���κα���ͼ��</font>

<%		  Response.end
			end if
          set drs=nothing
         %>
</td>
</tr>
<% end if %>