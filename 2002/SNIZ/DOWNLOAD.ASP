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
<!--#INCLUDE FILE="inc_functions.asp" -->
<!--#INCLUDE FILE="inc_top.asp" -->
<h3>DOWNLOAD:</h3>

<ul>
<li><a href="download.asp?file_id=18"><% =strVersion %></a></li>
</ul>

<h3>NEW FEATURES:</h3>

<ul>
<li>Locking/Unlocking/Deleting Members</li>
<li>More Private Forums Fixes</li>
<li>Private Forums Hidden Forums</li>
<li>Dates Fixed</li>
<li><b>Now converting |, ', . and ... to exact ASCII code. PROPERLY!</b></li>
<li>Fixed Forum Description bug</li>
<li>HTML can now be enabled Admin Section</li>
<li>Forum Code can now be disabled in the Admin Section</li>
<li>Lots of Tweaks</li>
<li>Re-Formated Reply/NewTopic Forms</li>
<li>Extended Error Checking in forums</li>
<li>Lots of Bug Fixes</li>
<li>Fixed Registration Bugs</li>
<li>doCode function Upgraded</li>
<li>Wicked Cool Stars</li>
<li>SQL Server Compliance Phase 3 complete (should be done, so please test and post results)</li>
<li>Several Bug Fixes Too</li>
<li>Some Bug Fixes, etc...</li>
</ul>

<h3>UPDATES:</h3>

<ul>
<li><b>FORUM_CONFIG</b>
<ul>
<li>ALTER : C_COPYRIGHT : C_STRCOPYRIGHT</li>
<li>ADD : C_STRRANKLEVEL0 : TEXT : 50</li>
<li>ADD : C_STRRANKLEVEL5 : TEXT : 50</li>
<li>ADD : C_STRRANKCOLORADMIN : TEXT : 50</li>
<li>ADD : C_STRRANKCOLORMOD : TEXT : 50</li>
<li>ADD : C_STRRANKCOLOR0 : TEXT : 50</li>
<li>ADD : C_STRRANKCOLOR1 : TEXT : 50</li>
<li>ADD : C_STRRANKCOLOR2 : TEXT : 50</li>
<li>ADD : C_STRRANKCOLOR3 : TEXT : 50</li>
<li>ADD : C_STRRANKCOLOR4 : TEXT : 50</li>
<li>ADD : C_STRRANKCOLOR5 : TEXT : 50</li>
<li>ADD : C_INTRANKLEVEL0 : NUMBER : INTEGER</li>
<li>ADD : C_INTRANKLEVEL1 : NUMBER : INTEGER</li>
<li>ADD : C_INTRANKLEVEL5 : NUMBER : INTEGER</li>
<li>ADD : C_STRALLOWFORUMCODE : NUMBER : BYTE : 1/0
</ul>
</li>
<li><b>FORUM_MEMBERS</b>
<ul>
<li>ADD : M_LASTHEREDATE : TEXT : 50</li>
<li>ALTER : M_SIG : TEXT : 255 </li>
</ul>
</li>
</ul>

<h3>注意：</h3>

<p>All <b>Yes/No</b> (including all TEXT fields that have "true" and "false" values) columns must be altered to be <b>Number</b> fields set to <b>byte</b> and as appropriate, set the default values to <b>1/0</b>.</p>

<!--#INCLUDE FILE="inc_footer.asp" -->
