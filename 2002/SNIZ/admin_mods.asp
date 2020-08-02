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
<!--#INCLUDE file="inc_functions.asp" -->
<!--#INCLUDE file="inc_top.asp" -->

<%

	Dim objDict
	Set objDict = Server.CreateObject("Scripting.Dictionary")
	set objRec = my_Conn.execute("SELECT M_CODE, M_VALUE FROM " & strTablePrefix & "MODS WHERE M_NAME = 'HModEnable'")

	while not objRec.EOF 
	objDict.Add objRec.Fields.Item("m_code").Value, objRec.Fields.Item("m_value").Value
	objRec.moveNext
	wend 

	if Request.Form("Method_Type") = "Write_Configuration" then 
	
		Err_Msg = ""
	
		if Err_Msg = "" then

			'## Forum_SQL
			objDict.Item("Attachment") = Request.Form("strAllowUploads")
			objDict.Item("Pmessages")= Request.Form("strPmessages")
			objDict.Item("UserFields")= Request.Form("strUserFields")
			objDict.Item("PollMentor")= Request.Form("strPollMentor")
			objDict.Item("SideMenu")= Request.Form("strSideMenu")
			objDict.Item("NewMemberPM")= Request.Form("strOnOff")
			objDict.Item("imageURLPath")= Request.Form("fstrImageURL")
			a = objDict.Keys
			b = objDict.Items
   			For i = 0 To objDict.Count -1 ' Iterate the array.
			strSql = "UPDATE " & strTablePrefix & "MODS "
			strSql = strSql & " SET M_VALUE ='" & b(i) & "'"
			strSQL = strSql & " WHERE M_NAME = 'HModEnable' AND M_CODE = '" & a(i) & "' "

			my_Conn.Execute (strSql)
   			Next
			Application(strCookieURL & "ConfigLoaded") = ""

%>
<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">设置已经更新!</font></p>


<p align="center"><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">修改完成！</font></p>

<%		else %>
	<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">你输入的资料有问题或没有填写完整</font></p>

	<table align=center border=0>
	  <tr>
	    <td><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
		<ul>
		<% =Err_Msg %>
		</ul>
	    </font></td>
	  </tr>
	</table>

	<p align=center><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>><a href="JavaScript:history.go(-1)">请返回重新输入</a></font></p>
<%		end if 
	else
%>
<form action="admin_mods.asp" method="post" id="Form1" name="Form1">
<input type="hidden" name="Method_Type" value="Write_Configuration">
<table border="0" cellspacing="1" cellpadding="3" align=center bgcolor="<% =strPopUpBorderColor %>">
  <tr valign="top">
    <td bgcolor="<% =strHeadCellColor %>" colspan="2"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>"><b>HuwR Mod Configuration</b></font></td>
  </tr>
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">上传文件:&nbsp;</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    开启：<input type="radio" class="radio" name="strAllowUploads" value="1" <% if intAllowUploads <> 0 then Response.Write("checked") %>> 
    关闭：<input type="radio" class="radio" name="strAllowUploads" value="0" <% if intAllowUploads = 0 then Response.Write("checked") %>>

    </td>
  </tr>
    <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">悄悄话信息:&nbsp;</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    开启：<input type="radio" class="radio" name="strPMessages" value="1" <% if intPMessages <> 0 then Response.Write("checked") %>> 
    关闭：<input type="radio" class="radio" name="strPMessages" value="0" <% if intPMessages = 0 then Response.Write("checked") %>>

    </td>
  </tr>
  </tr>
    <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">额外会员资料栏:&nbsp;</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    开启：<input type="radio" class="radio" name="strUserFields" value="1" <% if intUserFields <> 0 then Response.Write("checked") %>> 
    关闭：<input type="radio" class="radio" name="strUserFields" value="0" <% if intUserFields = 0 then Response.Write("checked") %>>

    </td>
  </tr>
  </tr>
    <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">论坛投票:&nbsp;</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    开启：<input type="radio" class="radio" name="strPollMentor" value="1" <% if intPollMentor <> 0 then Response.Write("checked") %>> 
    关闭：<input type="radio" class="radio" name="strPollMentor" value="0" <% if intPollMentor = 0 then Response.Write("checked") %>>

    </td>
  </tr>
    <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">快捷功能菜单:&nbsp;</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    开启：<input type="radio" class="radio" name="strSideMenu" value="1" <% if intSideMenu <> 0 then Response.Write("checked") %>> 
    关闭：<input type="radio" class="radio" name="strSideMenu" value="0" <% if intSideMenu = 0 then Response.Write("checked") %>>

    </td>
  </tr>
   <tr valign="top">
       <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">发送悄悄话给新会员:&nbsp;</font></td>
		   <td bgColor="<% =strPopUpTableColor %>">
		   开启：<input type="radio" class="radio" name="strOnOff" value="1" <% if intNewMemberPM <> 0 then Response.Write("checked") %>> 
		   关闭：<input type="radio" class="radio" name="strOnOff" value="0" <% if intNewMemberPM = 0 then Response.Write("checked") %>>
	   </td>
   </tr>
  
  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" align="right"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">Image图片路径:&nbsp;</font></td>
    <td bgColor="<% =strPopUpTableColor %>">
    <input type="text" name="fstrImageURL" value="<%=strImageURL %>" >
    </td>
  </tr>

  <tr valign="top">
    <td bgColor="<% =strPopUpTableColor %>" colspan="2" align="center"><input type="submit" value="提交设置" id="submit1" name="submit1"> <input type="reset" value="重新设置" id="reset1" name="reset1"></td>
  </tr>
</table>
</form>
</font>
<%	end if %>
<div align="center">    <p><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><a href="admin_home.asp">返回管理中心</A></font></p></div>
    </font>
    </center>
    </td>
  </tr>
</table>

</BODY>
</HTML>
<!--#include file="inc_footer.asp"-->
