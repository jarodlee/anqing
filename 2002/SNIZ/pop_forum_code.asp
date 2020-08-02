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

<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td bgcolor="<% =strTableBorderColor %>">
<table border="0" width="100%" cellspacing="1" cellpadding="4">
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a name="format"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" ><b>怎样使用本论坛专用代码...</b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>">
    <blockquote>
      <p><b>粗体字:</b> 将文本字体加粗显示..用法： [b] 文字 [/b] .&nbsp; <i>示例:</i> 这是一个 <b>[b]</b>粗体字<b>[/b]</b>. = 这是一个 <b>粗体字</b> .</p>

      <p><i>斜体字:</i> 将文本字体设置为斜体..用法 [i] 文字 [/i] .&nbsp; <i>示例:</i> 这是一个 <b>[i]</b>斜体字<b>[/i]</b> . = 这是一个 <i>斜体字</i> .</p>

      <p><u>下划线:</u> 将指定文字加上下划线..用法 [u] 文字 [/u]. <i>示例:</i> 这是一个 <b>[u]</b>下划线<b>[/u]</b> . =  这是一个 <u>下划线</u> .</p>

      <p>Aligning Text Left:<br>
        Enclose your text with [left] and [/left]
      </p>

      <p>Aligning Text Center:<br>
        Enclose your text with [center] and [/center]
      </p>

      <p>Aligning Text Right:<br>
        Enclose your text with [right] and [/right]
      </p>

      <p>Striking Text:<br>
        Enclose your text with [s] and [/s]<br>
        Example: <b>[s]</b>mistake<b>[/s]</b> = <s>mistake</s>
      </p>

      <p>&nbsp; </p>

      <p><b>Font Colors:</b><br>
        Enclose your text with [<i>fontcolor</i>] and [/<i>fontcolor</i>] <br>
        Example: <b>[red]</b>Text<b>[/red]</b> = <font color="red">Text</font id=red><br>
        Example: <b>[blue]</b>Text<b>[/blue]</b> = <font color="blue">Text</font id=blue><br>
        Example: <b>[pink]</b>Text<b>[/pink]</b> = <font color="pink">Text</font id=pink><br>
        Example: <b>[brown]</b>Text<b>[/brown]</b> = <font color="brown">Text</font id=brown><br>
        Example: <b>[black]</b>Text<b>[/black]</b> = <font color="black">Text</font id=black><br>
        Example: <b>[orange]</b>Text<b>[/orange]</b> = <font color="orange">Text</font id=orange><br>
        Example: <b>[violet]</b>Text<b>[/violet]</b> = <font color="violet">Text</font id=violet><br>
        Example: <b>[yellow]</b>Text<b>[/yellow]</b> = <font color="yellow">Text</font id=yellow><br>
        Example: <b>[green]</b>Text<b>[/green]</b> = <font color="green">Text</font id=green><br>
        Example: <b>[gold]</b>Text<b>[/gold]</b> = <font color="gold">Text</font id=gold><br>
        Example: <b>[white]</b>Text<b>[/white]</b> = <font color="white">Text</font id=white><br>
        Example: <b>[purple]</b>Text<b>[/purple]</b> = <font color="purple">Text</font id=purple>
      </p>

      <p>&nbsp; </p>

      <p><b>Headings:</b><br>
        Enclose your text with [h<i>number</i>] and [/h<i>n</i>]<br>
        <table border=0>
          <tr>
            <td><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
            Example: <b>[h1]</b>Text<b>[/h1]</b> =
            </font></td>
            <td><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
            <h1>Text</h1>
            </font></td>
          </tr>
          <tr>
            <td><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
            Example: <b>[h2]</b>Text<b>[/h2]</b> =
            </font></td>
            <td><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
            <h2>Text</h2>
            </font></td>
          <tr>
            <td><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
            Example: <b>[h3]</b>Text<b>[/h3]</b> =
            </font></td>
            <td><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
            <h3>Text</h3>
            </font></td>
          </tr>
          <tr>
            <td><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
            Example: <b>[h4]</b>Text<b>[/h4]</b> =
            </font></td>
            <td><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
            <h4>Text</h4>
            </font></td>
          </tr>
          <tr>
            <td><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
            Example: <b>[h5]</b>Text<b>[/h5]</b> =
            </font></td>
            <td><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
            <h5>Text</h5>
            </font></td>
          </tr>
          <tr>
            <td><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
            Example: <b>[h6]</b>Text<b>[/h6]</b> =
            </font></td>
            <td><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
            <h6>Text</h6>
            </font></td>
          </tr>
        </table>
      </p>

      <p>&nbsp; </p>

      <p><b>Font Sizes:</b><br>
        Example: <b>[size=1]</b>text<b>[/size=1]</b> = <font size=1>Text</font id=size1><br>
        Example: <b>[size=2]</b>text<b>[/size=2]</b> = <font size=2>Text</font id=size2><br>
        Example: <b>[size=3]</b>text<b>[/size=3]</b> = <font size=3>Text</font id=size3><br>
        Example: <b>[size=4]</b>text<b>[/size=4]</b> = <font size=4>Text</font id=size4><br>
        Example: <b>[size=5]</b>text<b>[/size=5]</b> = <font size=5>Text</font id=size5><br>
        Example: <b>[size=6]</b>text<b>[/size=6]</b> = <font size=6>Text</font id=size6>
      </p>

      <p>&nbsp; </p>

      <p>Bulleted List: <b>[list]</b> and <b>[/list]</b>, and items in list with <b>[*]</b> and <b>[/*]</b>.</p>

      <p>Ordered Alpha List: <b>[list=a]</b> and <b>[/list=1]</b>, and items in list with <b>[*]</b> and <b>[/*]</b>.</p>

      <p>Ordered Number List: <b>[list=1]</b> and <b>[/list=1]</b>, and items in list with <b>[*]</b> and <b>[/*]</b>.</p>

      <p>Code: enclose your text with <b>[code]</b> and <b>[/code]</b>.</p>

      <p>Quote: enclose your text with <b>[quote]</b> and <b>[/quote]</b>.</p>

<%	if (strIMGInPosts = "1") then %>
      <p>Images: enclose the address with <b>[img]</b> and <b>[/img]</b>.</p>
<%	end if %>
    </blockquote></font>
    </td>
  </tr>
</table>
    </td>
  </tr>
</table>

<!--#INCLUDE FILE="inc_footer_short.asp" -->
