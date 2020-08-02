newtopic.ASP文件的调用方法<BR><BR>



<p>演示,论坛主题调用</p>
在页面插入的代码:<p>&lt;script src=<font color="#800000">http://www.artbbs.net/news/</font>newtopic.asp?boardid=all&amp;tlen=16&amp;n=10&amp;sdate=1&amp;orders=2&amp;info=1&amp;action=<font color="#FF6699">1</font>&gt;&lt;/script></p>

<script src=newtopic.asp?boardid=all&tlen=16&n=10&sdate=1&orders=2&info=1&action=1></script>

相关参数<br>
boardid：版面id，全部为all<br>
tlen：标题长度<br>
n：显示多少个标题<br>
sdate：查询多少天内帖子，1为当天<br>
orders：排序方法，1为按照点击(最热帖)，2为按照时间(最新帖)<br>
info：1为显示发表时间和用户，2为显示发表时间，3为显示发表用户，4为显示发表用户和点击数，5为显示点击数，6为显示发表日期和用户，7为显示发表日期，0为不显示<br>
action:   1: 显示主题  2:显示精华主题 3:显示主题或回复<BR>
  

<br>
<BR><BR>
============================================================================<BR><BR>
tongji.asp文件的调用方法:<BR><BR>


以下调用:<BR>
1: orders=1 调用发帖TOP用户
2: orders=2 论坛统计

<br>
发帖前5名用户
<BR>
在页面插入的代码:<p>&lt;script src=<font color="#800000">http://www.artbbs.net/news/</font>tongji.asp?orders=2&amp;N=10>&lt;/script></p>

<script src=tongji.asp?orders=2&N=10></script>
<br>
相关参数<br>
n：显示多少个名<br>




<br>
论坛统计
<BR>
在页面插入的代码:<p>&lt;script src=<font color="#800000">http://www.artbbs.net/news/</font>tongji.asp?orders=1>&lt;/script></p>

<script src=tongji.asp?orders=1></script>