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
<script language="Javascript">
<!-- hide

function insertsmilie(smilieface){

	PostTopic.Message.value+=smilieface;
}
// -->
</script>
<tr>
<td bgColor="<% =strPopUpTableColor %>" align=right rowspan=2 valign=top>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">格式：</font>
</td>
<td bgColor="<% =strPopUpTableColor %>" align=left>
<a href="Javascript:bold();"><img src="<%=strImageURL %>icon_editor_bold.gif" width="22" height="22" alt="粗体" border="0"></a><a href="Javascript:italicize();"><img src="<%=strImageURL %>icon_editor_italicize.gif" width="23" height="22" alt="斜体" border="0"></a><a href="Javascript:underline();"><img src="<%=strImageURL %>icon_editor_underline.gif" width="23" height="22" alt="下划线" border="0"></a>
<a href="Javascript:center();"><img src="<%=strImageURL %>icon_editor_center.gif" width="22" height="22" alt="居中" border="0"></a>
<a href="Javascript:hyperlink();"><img src="<%=strImageURL %>icon_editor_url.gif" width="22" height="22" alt="插入超级连接" border="0"></a><a href="Javascript:email();"><img src="<%=strImageURL %>icon_editor_email.gif" width="23" height="22" alt="插入邮件地址" border="0"></a><a href="Javascript:image();"><img src="<%=strImageURL %>icon_editor_image.gif" width="23" height="22" alt="插入图片" border="0"></a>
<a href="Javascript:showcode();"><img src="<%=strImageURL %>icon_editor_code.gif" width="22" height="22" alt="插入代码" border="0"></a><a href="Javascript:quote();"><img src="<%=strImageURL %>icon_editor_quote.gif" width="23" height="22" alt="插入引用" border="0"></a><a href="Javascript:list();"><img src="<%=strImageURL %>icon_editor_list.gif" width="23" height="22" alt="插入段落" border="0"></a>
</td>
</tr><tr>
<td bgColor="<% =strPopUpTableColor %>" align=left>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
    <select name="font" onChange="showfont(this.options[this.selectedIndex].value)">
	<option value="宋体" selected>宋体</option>
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
	<option value="black" selected>黑色</option>
	<option value="red">红色</option>
	<option value="yellow">黄色</option>
	<option value="pink">粉红色</option>
	<option value="green">绿色</option>
	<option value="orange">橘色</option>
	<option value="purple">紫色</option>
	<option value="blue">蓝色</option>
	<option value="beige">米黄色</option>
	<option value="brown">棕色</option>
	<option value="teal">蓝绿色</option>
	<option value="navy">深蓝色</option>
	<option value="maroon">褐紫红色</option>
	<option value="limeGreen">淡黄绿色</option>
</select></td>
</tr>
<% if lcase(strIcons) = "1" then %>
<tr><td bgColor="<% =strPopUpTableColor %>" align=right rowspan=1 valign=top>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">表情符号：</font><br>
<a href="JavaScript:openWindow5('pop_icon_legend.asp')">更多表情符号</a>
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
          <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">没有任何表情图案</font>

<%		  Response.end
			end if
          set drs=nothing
         %>
</td>
</tr>
<% end if %>