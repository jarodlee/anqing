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
<!--#include file="config.asp"-->
<!--#include file="inc_functions.asp"-->
<!--#include file="inc_top_short.asp"-->
<% mLev = cint(ChkUser2(strDBNTUserName, Request.Cookies(strUniqueID & "User")("Pword"))) %>
<br><br><br>
<div align="center"> <FORM enctype="multipart/form-data" action="admin_UploadEngine.asp" method=POST name="fileup">
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
   <TABLE width="80%" border=0 cellspacing=1 cellpadding=4 align="center">
    <TR align="center" bgColor="<% =strPopUpTableColor %>">

     <TD align="center">
	 请选择要上传的文件
      <INPUT name="Folder" type="hidden" value="<%= strDBNTUsername %>" >
      <input name="REPLY_ID" type="hidden" value="<% =Request("REPLY_ID") %>">
	  <input name="TOPIC_ID" type="hidden" value="<% =Request("TOPIC_ID") %>">
	  <input name="MEMBER_ID" type="hidden" value="<% =Request("MEMBER_ID") %>">
    </TD>
    </TR>
    <TR>

     <TD align="center">
      <INPUT name="file1" type="FILE" size=20 value="">
     </TD>
	 <td></td>
    </TR>
	<TR>
     <TD align="center"><br>
      <INPUT name="Upload" type=SUBMIT value="开始上传" >
     </TD>
	 <td></td>
    </TR>
	<% If mlev = 4 then %>
		<TR>
     <TD align="center"><br>
      存放目录 <INPUT name="uploaddir" type="text" value="tools" >
     </TD>
	 <td></td>
    </TR>
<% End If %>
  </FORM>
   </TABLE>
   </font>
<!--#include file="inc_footer_short.asp"-->

