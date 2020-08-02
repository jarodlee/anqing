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
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td bgcolor="<% =strTableBorderColor %>">
<table border="0" width="100%" cellspacing="1" cellpadding="4">
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>"><b>论坛常见问题集</b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>">
    <p><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>">
    <ul>
    <li><a href="#register">我必须注册吗？</a></li>
<% if (stricons = "1") then %>
    <li><a href="#smilies">我要如何使用表情符号还有贴图？</a></li>
<% end if %>
    <li><a href="#hyperlink">我能在文章中加入超级连接吗？</a></li>
    <li><a href="#format">我要如何更改文字输出的格式？</a></li>
    <li><a href="#mods">版主在论坛中是做什么的？</a></li>
    <li><a href="#profile">我要如何修改我的注册和资料？</a></li>
    <li><a href="#cookies">有使用Cookie纪录吗？</a></li>
    <li><a href="#activetopics">什么是最新文章？</a></li>
    <li><a href="#edit">我能够修改我的文章吗？</a></li>
    <li><a href="#attach">我可以夹带附件吗？</a></li>
    <li><a href="#search">我可以搜索文章吗？</a></li>
    <li><a href="#EditProfile">我可以修改个人资料吗？</a></li>
    <li><a href="#signature">我能在我的文章下面附上签名吗？</a></li>
<!--    <li><A href="#announce">What are announcements?</A></li> -->
<% if (strBadWordFilter = "1") then %>
    <li><a href="#censor">有任何审查制度吗？</a></li>
<% end if %>
<% if (strEmail = "1") then %>
    <li><a href="#pw">如果忘了帐号或密码我该怎么办？</a></li>
<% end if %>
<% if (strEmail = "1") then %>
    <li><a href="#notify">假如有人回复我的主题可以用电子邮件通知我吗？</a></li>
<% end if %>
    <li><a href="#COPPA">什么是COPPA？</a></li>
    <!-- <%'## Added%> -->
    <%
    strsql = "SELECT F_ID, F_FAQ_QUESTION, F_FAQ_ANSWER FROM " & strTablePrefix & "FAQ ORDER BY F_ID DESC"
    set frs = my_conn.execute(strsql)
    	if frs.eof then
        	'## Do nothing
        else
        	Do until frs.eof
            
            response.write("<li><a href=""#faqid" & frs("F_ID") & """>" & frs("F_FAQ_QUESTION") & "</a></li>" & vbcrlf)
            
            frs.movenext
            Loop
        End if
    set frs = nothing
    %>
    <!-- <%'## End Added%> -->
    <li><a href="#getforum">能在哪里拿到本论坛的原代码呢？</a></li>
    <li><a href="mailto:<% =strSender %>">无法从这获得你想要的答案？请来信询问。</a></li>
    </ul>
    </font></p>
    </td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" align="right" border="0"></a><a name="register"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>"><b>注册</b></font></td>
  </tr>
  <tr>
      <td bgcolor="<% =strForumCellColor %>">
      <p><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
      如果没有注册的话，你可以自由的查看论坛中的任何文章，但如果你想要参与讨论的话， 
            就必须在论坛中注册。注册填写的资料很少，只要花费你一点点的时间而已，你只要填入会员名称和电子邮件地址
            以及一些相关资料，并设定好密码，就可以自由参与讨论，<% =strForumTitle %>保证不会将你的 
            注册资料提供给第三者使用。如果有任何使用上的疑问，请尽快反应，会第一时间处理。</font></p>
      
      </td>
  </tr>
<% If (strIcons = "1") then %>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" align="right" border="0"></a><a name="smilies"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>"><b>表情符号</b></font></td>
  </tr>
  <tr>
      <td bgcolor="<% =strForumCellColor %>">
      <p><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
      你或许在电子邮件或其他论坛看过别人使用表情符号，这是指一些代表情绪的键盘符号， 像是一个笑脸<img border="0" src="<%=strImageURL %>icon_smile.gif">或是难过的脸<img border="0" src="<%=strImageURL %>icon_smile_sad.gif">。这个讨论区系统可以自动将这些符号转换成图形档，譬如说文章中的这个符号 [:D] 会自动转换成这个图形<img border="0" src="<%=strImageURL %>icon_smile_big.gif">所有的表情图案都要记得加上[]中括号，下面就是<% =strForumTitle %>所有提供的表情图案，如果你有更丰富的表情图案建议，请主动提供：<br>
      <table border="0" align="center" cellpadding="5">
        <tr valign="top">
          <td>
          <table border="0" align="center">
          <%
          strsql = "select smile_name, smile_url, smile_code from "&strTablePrefix&"smiles"
          set drs=my_conn.execute(strsql)
          Do until drs.eof
          	smile_name=drs("smile_name")
            smile_url= drs("smile_url")
            smile_code=drs("smile_code")
            %>
              <tr>
                      <td bgcolor="<% =strForumCellColor %>"><img border="0" hspace="10" src="<%=strImageURL %><%=smile_url%>"></td>
                      <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><%=smile_name%></font></td>
                      <td bgcolor="<% =strForumCellColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">[<%=smile_code%>]</font></td>
              </tr>
              <%
              drs.movenext
              loop
              set drs=nothing
              %>
      </table>
          </td>
        </tr>
      </table>
      </font></p>
      </td>
  </tr>
<% end if %>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" align="right" border="0"></a><a name="hyperlink"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>"><b>在文章中使用超级连接</b></font></td>
  </tr>
  <tr>
      <td bgcolor="<% =strForumCellColor %>"><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
      <p>你可以在文章中轻松的加入超连结。</p>

      <p>你可以直接输入完整的网址（<% =strForumURL %>），系统会自行转成网址连接（<a href="<% =strForumURL %>" target="_blank"><% =strForumURL %></a>）！</p>
      
      <p>注意！一定要在网址前加上 <b>http://</b>, <b>https://</b> 或 <b>file://</b></p>

      <p>你也可以在文章中直接输入电子邮件信箱位址。<br>
      <blockquote>
			<b>例如：</b><br>
			<b><% =strSender %></b><br>
			<b>会变成：</b><br>
			<a href="mailto:<% =strSender %>"><% =strSender %></a>.</p>
      </blockquote>
      
      <p>另一种使用超级连接的方式是使用下面的 <b>[url]</b>网址<b>[/url]</b> 标签</p>
	  <blockquote>
              <b>例如：</b><br>
              <b>[url]</b><% =strForumURL %><b>[/url]</b> 回到论坛首页！<br>
              <b>会变成：</b><br>
              <a href="<% =strForumURL %>"><% =strForumURL %></a> 回到论坛首页！
      </blockquote></p>
	  <p> 
      <p>如果你使用以下格式：<b>[url=&quot;</b>连接<b>&quot;]</b>描述<b>[/url]</b> 你可以加入一个连接的名称或描述。</p>
      <blockquote>
              <b>例如：</b><br>
              连结到 <b>[url=&quot;<% =strForumURL %>&quot;]</b><% =strForumTitle %><b>[/url]</b><br>
              <b>会变成：</b><br>
              连接到 <a href="<% =strForumURL %>"><% =strForumTitle %></a>
      </blockquote>
      <blockquote>
              <b>例如：</b><br>
              如果你有任何疑问，<b>[url=&quot;<% =strSender %>&quot;]</b>欢迎来信询问<b>[/url]</b><br>
              <b>会变成：</b><br>
              如果你有任何疑问，<a href="mailto:<% =strSender %>">欢迎来信询问</a>
      </blockquote>
      <br> 
<%	if (strIMGInPosts = "1") then %>
      你可以加入图片连接，用 <b>[url="</b>连接<b>"]</b>描述<b>[/url]</b> 和 <b>[img]</b>图片地址<b>[/img]</b> 标签<br>
      <blockquote>
              <b>例如：</b><br>
              <b>[url=&quot;<% =strForumURL %>&quot;][img]</b><% =strForumURL %><%=strTitleImage%> <b>[/img][/url]</b><br>
              <b>会变成：</b><br>
              <a href="<% =strForumURL %>"><img src="<%=strTitleImage%>" target="_new" border="0"></a>
      </blockquote>
      </p>
<%	end if %>

<!--
      To do so place &quot;[url=&quot; in front of your url and &quot;]&quot; after 
      your url.&nbsp; Then type your description and end it with a &quot;[/url]&quot;.</p>
      <blockquote>
              <p><i>范例：</i><br>
              <b>[url=<% =strForumURL %>]</b><% =strForumTitle %><b>[/url]</b> takes you home!</p>
              
              <p><i>会变成：</i><br>
              <a href="<% =strForumURL %>"><% =strForumTitle %></a> takes you home!</p>
      </blockquote>
-->
      </font></td>
  </tr>
<% if strAllowForumCode = "1" then %>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" align="right" border="0"></a><a name="format"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>"><b>文字要如何使用粗体、斜体以及引用回复等功能？</b></font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strForumCellColor %>">
    <p><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
    论坛提供了数种文字格式可以让你改变文字的输出效果， 下面示例就是讨论区提供的文字输出效果：</p>
    <blockquote>
      <p><b>粗体</b>：在 [b] 和 [/b] 之间输入文字。 <i>示例：</i>这是<b>[b]</b>粗体<b>[/b]</b>文字。 = 这是<b>粗体</b>文字。</p>
      <p><i>斜体</i>：在 [i] 和 [/i] 之间输入文字。 <i>示例：</i>这是<b>[i]</b>斜体<b>[/i]</b>文字。 = 这是<i>斜体</i>文字。</p>
      <p><u>下划线</u>：在 [u] 和 [/u] 之间输入文字。 <i>示例：</i>这是<b>[u]</b>下划线<b>[/u]</b>文字。 = 这是<u>下划线</u>文字。</p>

      <p>文字靠左对齐：<br>
        在 [left] 和 [/left] 之间输入文字
      </p>

      <p>文字居中对齐：<br>
        在 [center] 和 [/center] 之间输入文字
      </p>

      <p>文字靠右对齐：<br>
        在 [right] 和 [/right] 之间输入文字
      </p>

      <p>删除文字：<br>
        在 [s] 和 [/s] 之间输入文字 <i>示例：</i> <b>[s]</b>删除<b>[/s]</b> = <s>删除</s>
      </p>

      <p>&nbsp; </p>

      <p><b>文字颜色：</b><br>
        在 [<i>文字颜色</i>] 和 [/<i>文字颜色</i>] 之间输入文字<br>
        <i>示例：</i> <b>[red]</b>文字<b>[/red]</b> = <font color="red">文字</font id=red><br>
        <i>示例：</i> <b>[blue]</b>文字<b>[/blue]</b> = <font color="blue">文字</font id=blue><br>
        <i>示例：</i> <b>[pink]</b>文字<b>[/pink]</b> = <font color="pink">文字</font id=pink><br>
        <i>示例：</i> <b>[brown]</b>文字<b>[/brown]</b> = <font color="brown">文字</font id=brown><br>
        <i>示例：</i> <b>[black]</b>文字<b>[/black]</b> = <font color="black">文字</font id=black><br>
        <i>示例：</i> <b>[orange]</b>文字<b>[/orange]</b> = <font color="orange">文字</font id=orange><br>
        <i>示例：</i> <b>[violet]</b>文字<b>[/violet]</b> = <font color="violet">文字</font id=violet><br>
        <i>示例：</i> <b>[yellow]</b>文字<b>[/yellow]</b> = <font color="yellow">文字</font id=yellow><br>
        <i>示例：</i> <b>[green]</b>文字<b>[/green]</b> = <font color="green">文字</font id=green><br>
        <i>示例：</i> <b>[gold]</b>文字<b>[/gold]</b> = <font color="gold">文字</font id=gold><br>
        <i>示例：</i> <b>[white]</b>文字<b>[/white]</b> = <font color="white">文字</font id=white><br>
        <i>示例：</i> <b>[purple]</b>文字<b>[/purple]</b> = <font color="purple">文字</font id=purple>
      </p>

      <p>&nbsp; </p>

      <p><b>标题文字：</b><br>
        在 [h] 和 [/h] 之间输入文字<br>
        <table border=0>
          <tr>
            <td><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
            <i>示例：</i> <b>[h1]</b>文字<b>[/h1]</b> =
            </font></td>
            <td><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
            <h1>文字</h1>
            </font></td>
          </tr>
          <tr>
            <td><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
            <i>示例：</i> <b>[h2]</b>文字<b>[/h2]</b> =
            </font></td>
            <td><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
            <h2>文字</h2>
            </font></td>
          <tr>
            <td><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
            <i>示例：</i> <b>[h3]</b>文字<b>[/h3]</b> =
            </font></td>
            <td><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
            <h3>文字</h3>
            </font></td>
          </tr>
          <tr>
            <td><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
            <i>示例：</i> <b>[h4]</b>文字<b>[/h4]</b> =
            </font></td>
            <td><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
            <h4>文字</h4>
            </font></td>
          </tr>
          <tr>
            <td><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
            <i>示例：</i> <b>[h5]</b>文字<b>[/h5]</b> =
            </font></td>
            <td><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
            <h5>文字</h5>
            </font></td>
          </tr>
          <tr>
            <td><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
            <i>示例：</i> <b>[h6]</b>文字<b>[/h6]</b> =
            </font></td>
            <td><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
            <h6>文字</h6>
            </font></td>
          </tr>
        </table>
      </p>

      <p>&nbsp; </p>

      <p><b>字体大小：</b><br>
        <i>示例：</i> <b>[size=1]</b>文字<b>[/size=1]</b> = <font size=1>文字</font id=size1><br>
        <i>示例：</i> <b>[size=2]</b>文字<b>[/size=2]</b> = <font size=2>文字</font id=size2><br>
        <i>示例：</i> <b>[size=3]</b>文字<b>[/size=3]</b> = <font size=3>文字</font id=size3><br>
        <i>示例：</i> <b>[size=4]</b>文字<b>[/size=4]</b> = <font size=4>文字</font id=size4><br>
        <i>示例：</i> <b>[size=5]</b>文字<b>[/size=5]</b> = <font size=5>文字</font id=size5><br>
        <i>示例：</i> <b>[size=6]</b>文字<b>[/size=6]</b> = <font size=6>文字</font id=size6>
      </p>

      <p>&nbsp; </p>

      <p>项目列表：<b>[list]</b> 和 <b>[/list]</b>，列表中的内容可以在 <b>[*]</b> 和 <b>[/*]</b> 里。</p>

      <p>照字母排列列表：<b>[list=a]</b> 和 <b>[/list=a]</b>，列表中的内容可以在 <b>[*]</b> 和 <b>[/*]</b> 里。</p>

      <p>照数字标号列表：<b>[list=1]</b> 和 <b>[/list=1]</b>，列表中的内容可以在 <b>[*]</b> 和 <b>[/*]</b> 里。</p>

      <p>代码：在 <b>[code]</b> 和 <b>[/code]</b> 中输入完整或片段代码。</p>

      <p>引用：在 <b>[quote]</b> 和 <b>[/quote]</b> 中输入引用。</p>

<%	if (strIMGInPosts = "1") then %>
      <p>图片: 在 <b>[img]</b> 和 <b>[/img]</b>之间输入图片地址. 给图片加上连接可以使用 [url=""][img][/img][/url]</p>
<%	end if %>
    </blockquote></font>
    </td>
  </tr>
<% end if %>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" align="right" border="0"></a><a name="mods"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>"><b>版主</b></font></td>
  </tr>
  <tr>
      <td bgcolor="<% =strForumCellColor %>">
      <p><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
      版主主要的工作就是管理论坛中的文章，包括编辑、删除、锁定灌水以及人身攻击文章。 
<%	if (strShowModerators = "1") then %>
      如果有任何关於论坛的问题，请直接向论坛版主询问。
<%	end if %>
      </font></p>
      </td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" align="right" border="0"></a><a name="profile"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>"><b>更改个人资料</b></font></td>
  </tr>
  <tr>
      <td bgcolor="<% =strForumCellColor %>">
      <p><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
      更改个人资料很容易，只要点选上方或左方选单中的 &quot; 资料 &quot; ，并输入你的名称和密码，就可以直接更改你在论坛中的任何注册资料（除了名称）。
      </font></p>
      </td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" align="right" border="0"></a><a name="cookies"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>"><b>Cookies</b></font></td>
  </tr>
  <tr>
      <td bgcolor="<% =strForumCellColor %>">
      <p><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
      论坛可以设定使用Cookie来记录你的名称和密码，可以让你每次进入论坛免去重复打入名称密码的麻烦。只要你登陆后，下次再进入就不要打名称和密码了。如果你的浏览器没有打开或不支持Cookie功能，那么这个Cookie功能可能无法正常使用。
<% if (strNoCookies = "0") then %>
      <br><B>注意！如果有限制必须要注册才能观看的讨论区，你就要打开Cookie功能，包括浏览器的Cookie也要打开才能够发表新文章或是回复文章。</B>
<% end if %>
      </p>
      <p>你可以点选退出清除Cookie纪录的名称和密码，不过下次再进论坛就必须进行登陆。
      </font></p>
      </td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" align="right" border="0"></a><a name="activetopics"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>"><b>最新文章</b></font></td>
  </tr>
  <tr>
      <td bgcolor="<% =strForumCellColor %>">
      <font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
      <p>最新文章是由Cookie纪录提供的功能之一。当你点选了最新文章后，可以看到从你上次离开论坛（或是你指定的某段时间）之后有新文章发表的主题列表。</p>
      </font>
      </td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" align="right" border="0"></a><a name="Edit"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>"><b>编辑你的文章</b></font></td>
  </tr>
  <tr>
      <td bgcolor="<% =strForumCellColor %>">
      <p><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
      任何时候你都可以修改或删除你以前所留的文章，当你在检视你的文章时，可以在该则文章上方看到 （<img border="0" src="<%=strImageURL %>icon_edit_topic.gif" hspace="6"><img border="0" src="<%=strImageURL %>icon_delete_reply.gif" hspace="6">）
      这两个小图示，其中<img border="0" src="<%=strImageURL %>icon_edit_topic.gif" hspace="6">图标是用来修改留言，而<img border="0" src="<%=strImageURL %>icon_delete_reply.gif" hspace="6">图标则用来删除，除了你本人之外，只有管理者或板主才有权修改或删除你的留言，其他人是无权修改的。
<% if (strEditedByDate = "1") then %>
      编辑过后的文章，在最底下会标明重新修改日期和修改者。
<% end if %>
      </font></p>
      </td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" align="right" border="0"></a><a name="Attach"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>"><b>附加文件</b></font></td>
  </tr>
  <tr>
      <td bgcolor="<% =strForumCellColor %>">
      <p><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
      因为安全上的考虑，你无法在文章中附带任何形式的文件。但可以利用复制和贴上在你文章中转贴任何文字。
      </font></p>
      </td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" align="right" border="0"></a><a name="Search"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>"><b>搜索文章功能</b></font></td>
  </tr>
  <tr>
      <td bgcolor="<% =strForumCellColor %>">
      <p><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
      你可以搜索论坛里的全部文章。比方说你可以指定一个关键字，或是一个日期，甚至是一位作者……你都可以轻易的整理出符合您的搜索条件的文章。
      只要点选上方或左方选单中的 &quot; 搜索 &quot; 连结即可！
<!--      Note: announcements are not included in the search returns. -->
      </font></p>
      </td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" align="right" border="0"></a><a name="EditProfile"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>"><b>编辑个人资料</b></font></td>
  </tr>
  <tr>
      <td bgcolor="<% =strForumCellColor %>"><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
      <P>编辑个人资料很容易，只要点选上方或左方选单中的 &quot; 资料 &quot; ，并输入你的名称和密码，就可以直接更改你在论坛中的任何注册资料（除了名称）。</P>
      </font></td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" align="right" border="0"></a><a name="Signature"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>"><b>签名</b></font></td>
  </tr>
  <tr>
      <td bgcolor="<% =strForumCellColor %>"><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
      <p>你可以在每篇文章最后，让系统自动帮你加上你的签名。只要您在注册时有填上你独特的签名档，并在文章时勾选自动上加签名的功能即可。如果您在注册时没填，也可以在 &quot; 资料 &quot; 选单里随时加上去喔！</p>
      <p>注意: 签名档中不可以使用HTML语法，但可以使用表情图案。</p>
      </font></td>
  </tr>
<% if (strBadWordFilter = "1") then %>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" align="right" border="0"></a><a name="Censor"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>"><b>审查制度</b></font></td>
  </tr>
  <tr>
      <td bgcolor="<% =strForumCellColor %>">
      <p><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
      系统会自动过滤某些不良词语，比方说 FxxK 之类的。如果你在留言中打了这些字，将会发现系统自动将之转成 * 号。当然，如果你经常留这样的文字，最大的可能性是会被注销会员资格。
      </font></p>
      </td>
  </tr>
<% end if %>
<% if (strEmail = "1") then %>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" align="right" border="0"></a><a name="PW"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>"><b>忘了名称或密码</b></font></td>
  </tr>
  <tr>
      <td bgcolor="<% =strForumCellColor %>">
      <p><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
      忘了你的名称或密码都不是问题. 你可以利用 &quot; 忘记你的密码？&quot; 要求系统自动寄出你的名称及密码给你。当然，如果你的电子邮件地址是乱填的，那就无法收到了！
      </font></p>
      </td>
  </tr>
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" align="right" border="0"></a><a name="Notify"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>"><b>电子邮件通知</b></font></td>
  </tr>
  <tr>
      <td bgcolor="<% =strForumCellColor %>">
      <p><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
      当您发表了一个新主题之後，有时候会急为期待知道别人怎么样来回应你的主题。没问题！只要在发表新主题时，记得勾选 &quot; 当有人回复你的主题时以E-mail通知 &quot; 功能, 日后若有任何人回应你的主题, 系统就会自动寄 E-Mail 到你的信箱, 让你可以立即得知他人对你的主题的回应。
      </font></p>
      </td>
  </tr>
<% end if %>
    <!-- <%'## Added%> -->
    <%
    strsql = "SELECT F_ID, F_FAQ_QUESTION, F_FAQ_ANSWER FROM " & strTablePrefix & "FAQ ORDER BY F_ID DESC"
    set frs = my_conn.execute(strsql)
    strdisplay = ""
    	if frs.eof then
        	'## Do nothing
        else
        	Do until frs.eof
            
           ' response.write("<li><a href=""#faqid" & frs("F_ID") & """>" & frs("F_FAQ_QUESTION") & "</a></li>" & vbcrlf)
            strdisplay = strdisplay & "<tr><td bgcolor=""" & strCategoryCellColor & """><a href=""#top""><img src=""images/icon_go_up.gif"" align=""right"" border=""0""></a>"
			strdisplay = strdisplay & "<font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """>"
			strdisplay = strdisplay & "<b><a name=""faqid" & frs("F_ID") & """>" & frs("F_FAQ_QUESTION") & "</a></b></font></td></tr>"
            strdisplay = strdisplay & vbcrlf
            strdisplay = strdisplay & "<tr><td bgcolor=""" & strForumCellColor & """><font size=" & strDefaultFontSize & " face=""" & strDefaultFontFace & """ color=""" & strForumFontColor & """>" & vbcrlf
			strdisplay = strdisplay & "<p>" & frs("F_FAQ_ANSWER") & "</p></td></tr>"
            frs.movenext
            Loop
        End if
    if strdisplay <> "" then
    	response.write(strdisplay)
    End if
    set frs = nothing
    %>
    <!-- <%'## End Added%> -->
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" align="right" border="0"></a><a name="COPPA"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>"><b>什么是 COPPA</b></font></td>
  </tr>
  <tr>
      <td bgcolor="<% =strForumCellColor %>"><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
      <p>COPPA是『美国儿童网路隐私保护法（Children's Online Privacy Protection Act）』的简称。依其规定，商业网站必须先取得家长同意後，方能向13岁以下的儿童索取个人资料。同时，网站在未获孩童父母同意之前，也不得让儿童使用电子邮件、聊天室、或网络上任何可能与陌生人交换个人资讯的互动技术。简言之，父母们必须要参与儿童的网络活动，以确保自家孩子浏览的网站遵守COPPA规范。 
       -- 虽然它并非我国的法律条文，但本站亦将依该法精神来处理相关事宜。<a href="http://www.ftc.gov/bcp/conline/pubs/buspubs/coppa.htm" title="What is COPPA?">更详细的COPPA</a></p>
      </font></td>
  </tr>  
  <tr>
    <td bgcolor="<% =strCategoryCellColor %>"><a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" align="right" border="0"></a><a name="GetForum"></a><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>"><b>获得你自己的论坛</b></font></td>
  </tr>
  <tr>
      <td bgcolor="<% =strForumCellColor %>"><font size="<% =strDefaultFontSize %>" face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>">
            <p>你可以到 <a href="http://forum.snitz.com/" title="Link to Snitz Forums 2000 Homepage!">这个网站</a>下载最新的英文版论坛源代码。</p>
            </font></td>
  </tr>  
</table>
    </td>
  </tr>
</table>

<!--#INCLUDE FILE="inc_footer.asp" -->
