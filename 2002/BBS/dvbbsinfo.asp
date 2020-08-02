<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->

<%
boardid=request("boardid")
if not isnumeric(boardid) or boardid="" or boardid=0 then
boardid=1
end if


Stats="论坛变量信息"
call activeonline()
	call nav()
	call head_var(1,BoardDepth,0,0)


%>

<p>
</p>
<table cellpadding=6 cellspacing=1 align=center class=tableborder2>
    <tr>
      <td class=tablebody1>论坛变量信息</td>
    </tr>
    <tr>
      <td class=tablebody1>
<a href=?view=Setting&boardid=<%=boardid%> >常规设置信息</a> | 
<a href=?view=info&boardid=<%=boardid%> >基本信息</a> | 
<a href=?view=Group&boardid=<%=boardid%> >用户组权限设置</a> | 
<a href=?view=css&boardid=<%=boardid%> >关于颜色CSS的变量</a> | 
<a href=?view=pic&boardid=<%=boardid%> >论坛图片变量</a> | 
<a href=?view=boardpic&boardid=<%=boardid%> >论坛主体图标</a> | 
<a href=?view=TopicPic&boardid=<%=boardid%> >论坛帖子图标</a> | 
<a href=?view=userset&boardid=<%=boardid%> >用户设定</a> | 
<a href=?view=boardset&boardid=<%=boardid%> >分版设定</a>
</td>
    </tr>
  </table>

<%
if  request("view")="Setting" then
call Setting()
elseif request("view")="info" then
call info()
elseif request("view")="Group" then
call Group()
elseif request("view")="css" then
call css()
elseif request("view")="pic" then
call pic()
elseif request("view")="boardpic" then
call boardpic()
elseif request("view")="TopicPic" then
call TopicPic()
elseif request("view")="userset" then
call userset()
elseif request("view")="boardset" then
call boardset()
else
call Setting()
end if

call footer()

 sub Setting()%>
 <table cellpadding=6 cellspacing=1 align=center class=tableborder1 style="word-break:break-all;">
    <tr>
      <th height="24" colspan="2">Forum_Setting()  常规设置信息</th>
    </tr>
    <tr>
      <td class=tablebody1 colspan="2">
Forum_Setting()  常规设置信息  当前值设置如下<br>
0,300,1,1,0,0,1,1,20,0,0,20,10,16240,0,0,0,1,0,0,30,0,3000,0,1,0,500,0,1,0,1,1,1,100,1,0,0,1,0,10,0,0,1,10,0,0,1,1,1,0,0,0,0,0,0,1,1,0,1000,9,15,4,0,300,0,1,0,1,20,3,1,1,1
</td></tr>
<tr>
    <tr>
      <td width="70%" class=tablebody2 valign=top>
<li>forum_setting(0) 服务器时差    -ok
<li>forum_setting(1) 脚本超时时间    -ok
<li>forum_setting(2) 发送邮件组件  0－不支持 1－JMAIL 2－CDONTS  3－ASPEMAIL-ok
<li>forum_setting(3) 是否起用记录帖子阅读用户功能
<li>forum_setting(4) 首页模式   0-新版样式 1-旧版样式  
<li>forum_setting(5) 首页显示论坛深度-ok
<li>forum_setting(6) 用户头衔   0－关闭 1－打开  
<li>forum_setting(7) 头像上传   0－关闭 1－打开  
<li>forum_setting(8) 删除不活动用户时间 -ok   
<li>forum_setting(9) 
<li>forum_setting(10) 新短消息弹出窗口  0－否 1－是  
<li>forum_setting(11) 每页显示最多纪录    ==一般分页
<li>forum_setting(12) 
<li>forum_setting(13)
<li>forum_setting(14) 在线名单显示用户在线  0－关闭 1－打开  -ok
<li>forum_setting(15) 在线名单显示客人在线  0－关闭 1－打开  -ok
<li>forum_setting(16)    
<li>forum_setting(17) 
<li>forum_setting(18)    
<li>forum_setting(19) 防刷新机制   0－关闭 1－打开 -ok
<li>forum_setting(20) 浏览刷新时间间隔   -ok
<li>forum_setting(21) 论坛当前状态   -ok
<li>forum_setting(22) 同一IP注册间隔时间-ok   
<li>forum_setting(23) Email通知密码   0－关闭 1－打开 
<li>forum_setting(24) 一个Email只能注册一个帐号 0－关闭 1－打开 
<li>forum_setting(25) 注册需要管理员认证  0－关闭 1－打开 
<li>forum_setting(26) 允许同时在线数   -ok
<li>forum_setting(27) 
<li>forum_setting(28) 在线列表是否显示用户IP
<li>forum_setting(29) 是否显示过生日会员  0－否 1－是 
<li>forum_setting(30) 是否显示页面执行时间  0－否 1－是 
<li>forum_setting(31) 快速登陆位置   1－顶部 2－底部 “0”－不显示
<li>forum_setting(32) 是否开启论坛门派  0－否 1－是
<li>forum_setting(33) 在线资料列表显示当前位置
<li>forum_setting(34) 在线资料列表显示登陆和活动时间
<li>forum_setting(35) 在线资料列表显示浏览器和操作系统
<li>forum_setting(36) 在线资料列表显示来源（关闭后如果如果用户所属用户组设置了可以查看，则该用户可查看来源）
<li>forum_setting(37) 是否允许新用户注册
<li>forum_setting(38) 
<li>forum_setting(39) 
<li>forum_setting(40) 最短用户名长度
<li>forum_setting(41) 最长用户名长度  
<li>forum_setting(42) 允许个人签名
<li>forum_setting(43) 
<li>forum_setting(44) 作为热门话题的最低人气值  
<li>forum_setting(45)   
<li>forum_setting(46) 开启短信欢迎新注册用户  0－关闭 1－打开
<li>forum_setting(47) 发送注册信息邮件  0－关闭 1－打开
<li>forum_setting(48) 编辑过的帖子显示"由xxx于yyy编辑"的信息
<li>forum_setting(49) 管理员编辑后显示"由XXX编辑"的信息
<li>forum_setting(50) 等待"由XXX编辑"信息显示的时间
<li>forum_setting(51) 编辑帖子时限   
<li>forum_setting(52) 禁止的邮件地址
<li>forum_setting(53) 允许用户使用头像    --浏览帖子中设定
<li>forum_setting(54) 使用自定义头像的最少发帖数    
<li>forum_setting(55) 允许从其他站点上传头像
<li>forum_setting(56) 允许的最大头像文件大小
<li>forum_setting(57) 最大头像尺寸
<li>forum_setting(58) 查看图片选项
<li>forum_setting(59) 用户头衔最大长度
<li>forum_setting(60) 自定义头衔最少发帖数量限制
<li>forum_setting(61) 自定义头衔注册天数限制
<li>forum_setting(62) 自定义头衔上面两个条件加在一起限制
<li>forum_setting(63) 自定义头衔中要屏蔽的词语
<li>forum_setting(64) 防刷新功能有效的页面
<li>forum_setting(65) 用户签名是否开启UBB代码  0－关闭 1－打开 
<li>forum_setting(66) 用户签名是否开启HTML代码 0－关闭 1－打开 
<li>forum_setting(67) 用户是否开启贴图标签  0－关闭 1－打开 
<li>forum_setting(68) 用户排行列表个数   
<li>forum_setting(69) 是否定时开关论坛
<li>forum_setting(70) 定时开关论坛设置
<li>forum_setting(71) 
<li>forum_setting(72) 
</td>
      <td width="30%" class=tablebody1>
<table cellpadding=6 cellspacing=1 align=center width="100%" style="word-break:break-all;" valign=top>
<% for i=0 to 72%>
<tr ><td class=tableborder1 height=1  ></td></tr>
<tr>
      <td >Forum_Setting(<%=i%>) 当前值为：<%=Forum_Setting(i)%>

</td>
    </tr>
<% next %>
  </table>
</td>
    </tr>
    <tr>
      <th height="24" colspan="2"></th>
    </tr>
  </table>




<% end sub 
sub info()%>


 <table cellpadding=6 cellspacing=1 align=center class=tableborder1 style="word-break:break-all;">
    <tr>
      <th height="24" colspan="2">forum_info()  基本信息</th>
    </tr>
    <tr>
      <td class=tablebody1 colspan="2">
forum_info()   基本信息  默认信息设置
<li>
动网先锋论坛,http://www.dvbbs.net/,动网先锋,http://www.aspsky.net/,61.145.114.64,club@aspsky.net,images/LOGO.GIF,pic/,face/,北京时间


</td></tr>
<tr>
    <tr>
      <td width="30%" class=tablebody2 valign=top>
<li>Forum_info(0) 论坛名称
<li>Forum_info(1) 论坛的url
<li>Forum_info(2) 主页名称
<li>Forum_info(3) 主页URL
<li>Forum_info(4) SMTP Server地址
<li>Forum_info(5) 论坛管理员Email
<li>Forum_info(6) 论坛首页Logo地址
<li>Forum_info(7) 论坛图片目录
<li>Forum_info(8) 论坛表情目录
<li>Forum_info(9) 论坛所在时区
<li>Forum_info(10) 发贴表情目录
<li>Forum_info(11) 论坛头像目录



</td>
      <td width="70%" class=tablebody1>
<table cellpadding=6 cellspacing=1 align=center width="100%" style="word-break:break-all;" valign=top>
<% for i=0 to 11%>
<tr ><td class=tableborder1 height=1  ></td></tr>
<tr>
      <td >Forum_info(<%=i%>) 当前值为：<%=Forum_info(i)%>

</td>
    </tr>
<% next %>
  </table>
</td>
    </tr>
    <tr>
      <th height="24" colspan="2"></th>
    </tr>
  </table>

<% end sub 
sub Group()%>

 <table cellpadding=6 cellspacing=1 align=center class=tableborder1 style="word-break:break-all;">
    <tr>
      <th height="24" colspan="2">GroupSetting()[reGroupSetting()]  用户组权限设置</th>
    </tr>
    <tr>
      <td width="75%" class=tablebody2 valign=top>
GroupSetting()[reGroupSetting()]  用户组权限设置
<li>
<li>UserGroupID title
<li>1  管理员
<li>2  超级版主
<li>3  版主
<li>4  注册用户
<li>5  等待验证的(COPPA)会员
<li>6  等待邮件验证的会员
<li>7  未注册/未登录用户
<li>8  贵宾
<hr class=tableborder1>
<li>GroupSetting(0)  可以浏览论坛   0－否 1－是
<li>GroupSetting(1)  可以查看会员信息  0－否 1－是
<li>GroupSetting(2)  可以查看其他人发布的主题 0－否 1－是
<li>GroupSetting(3)  可以发布新主题   0－否 1－是
<li>GroupSetting(4)  可以回复自己的主题  0－否 1－是
<li>GroupSetting(5)  可以回复其他人的主题  0－否 1－是
<li>GroupSetting(6)  可以在论坛允许评分的时候参与评分0－否 1－是
<li>GroupSetting(7)  可以上传附件   0－否 1－是
<li>GroupSetting(8)  可以发布新投票   0－否 1－是
<li>GroupSetting(9)  可以参与投票   0－否 1－是
<li>GroupSetting(10) 可以编辑自己的帖子  0－否 1－是
<li>GroupSetting(11) 可以删除自己的帖子  0－否 1－是
<li>GroupSetting(12) 可以移动自己的帖子到其他论坛 0－否 1－是
<li>GroupSetting(13) 以打开/关闭自己发布的主题 0－否 1－是
<li>GroupSetting(14) 可以搜索论坛   0－否 1－是
<li>GroupSetting(15) 可以使用''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''发送本页给好友''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''功能 0－否 1－是
<li>GroupSetting(16) 可以修改个人资料  0－否 1－是
<li>GroupSetting(17) 可以发布小字报   0－否 1－是
<li>GroupSetting(18) 可以删除其它人帖子  0－否 1－是
<li>GroupSetting(19) 可以移动其它人帖子  0－否 1－是
<li>GroupSetting(20) 可以打开/关闭其它人帖子  0－否 1－是
<li>GroupSetting(21) 可以固顶/解除固顶帖子  0－否 1－是
<li>GroupSetting(22) 可以奖励/惩罚发贴用户  0－否 1－是
<li>GroupSetting(23) 可以编辑其它人帖子  0－否 1－是
<li>GroupSetting(24) 可以加入/解除精华帖子  0－否 1－是
<li>GroupSetting(25) 可以发布公告   0－否 1－是
<li>GroupSetting(26) 可以管理公告   0－否 1－是
<li>GroupSetting(27) 可以管理小字报   0－否 1－是
<li>GroupSetting(28) 可以锁定/屏蔽/解除锁定用户 0－否 1－是
<li>GroupSetting(29) 可以删除用户1－10天内所发帖子 0－否 1－是
<li>GroupSetting(30) 可以查看来访IP及来源  0－否 1－是
<li>GroupSetting(31) 可以限定IP来访   0－否 1－是
<li>GroupSetting(32) 可以发送短信   0－否 1－是
<li>GroupSetting(33) 最多发送用户  
<li>GroupSetting(34) 短信内容大小限制  
<li>GroupSetting(35) 信箱大小限制  
<li>GroupSetting(36) 是否有审核帖子的权限 0－否 1－是
<li>GroupSetting(37) 是否有进入隐含论坛的权限 0－否 1－是
<li>GroupSetting(38) 帖子中可以可以查看其它人签名 0－否 1－是
<li>GroupSetting(39) 可以浏览论坛事件  0－否 1－是
<li>GroupSetting(40) 最多上传文件个数  
<li>GroupSetting(41) 可以浏览精华帖子  0－否 1－是
<li>GroupSetting(42) 可以管理用户权限  0－否 1－是
<li>GroupSetting(43) 可以奖励/惩罚用户  0－否 1－是
<li>GroupSetting(44) 上传文件大小限制  
<li>GroupSetting(45) 可以批量删除帖子（前台） 0－否 1－是
<li>GroupSetting(46) 发布小字报所需金钱  
<li>GroupSetting(47) 参与评分所需金钱


</td>
      <td width="25%" class=tablebody1>
<table cellpadding=6 cellspacing=1 align=center width="100%" style="word-break:break-all;" valign=top>
<% for i=0 to 47%>
<tr ><td class=tableborder1 height=1  ></td></tr>
<tr>
      <td >GroupSetting(<%=i%>) 当前值为：<%=GroupSetting(i)%>

</td>
    </tr>
<% next %>
  </table>
</td>
    </tr>
    <tr>
      <th height="24" colspan="2"></th>
    </tr>
  </table>


<% end sub 
sub css()%>
 <table cellpadding=6 cellspacing=1 align=center class=tableborder1 style="word-break:break-all;">
    <tr>
      <th height="24" colspan="2">Forum_body()  关于颜色CSS的变量
</th>
    </tr>
    <tr>
      <td class=tablebody1 colspan="2">
Forum_body()  关于颜色CSS的变量

</td></tr>
<tr>
    <tr>
      <td width="30%" class=tablebody2 valign=top>
<li>Forum_body()  关于颜色CSS的变量
<li>Forum_body(0) 表格边框（体）CSS定义一 调用：Class=TableBorder1
<li>Forum_body(1) 表格边框（体）CSS定义二 调用：Class=TableBorder2
<li>Forum_body(2) 标题表格CSS定义一（深背景） 调用：th
<li>Forum_body(3) 标题表格CSS定义二（浅背景） 调用：Class=TableTitle2
<li>Forum_body(4) 表格体CSS定义一 调用：Class=TableBody1
<li>Forum_body(5) 表格体颜色二(1和2颜色在显示中穿插) 调用：Class=TableBody2
<li>Forum_body(6) 论坛连接总的CSS定义 
<li>Forum_body(7) 标题表格字体连接CSS定义（深背景） 调用：id=TableTitleLink
<li>Forum_body(8) 警告提醒语句的颜色 
<li>Forum_body(9) 显示帖子的时候，相关帖子，转发帖子，回复等的颜色 
<li>Forum_body(10) 论坛表格总的CSS定义 
<li>Forum_body(11) 论坛BODY标签 
<li>Forum_body(12) 表格宽度 
<li>Forum_body(13)  
<li>Forum_body(14) 首页连接颜色 
<li>Forum_body(15) 一般用户名称字体颜色 
<li>Forum_body(16) 一般用户名称上的光晕颜色 
<li>Forum_body(17) 版主名称字体颜色 
<li>Forum_body(18) 版主名称上的光晕颜色 
<li>Forum_body(19) 管理员名称字体颜色 
<li>Forum_body(20) 版主名称上的光晕颜色 
<li>Forum_body(21) 贵宾名称字体颜色 
<li>Forum_body(22) 贵宾名称上的光晕颜色 
<li>Forum_body(23) 顶部菜单表格CSS定义(Logo & Banner下方) 调用：Class=TopLighNav
<li>Forum_body(24) 顶部菜单表格CSS定义(Logo & Banner上方) 调用：Class=TopDarkNav
<li>Forum_body(25) Body的CSS定义 
<li>Forum_body(26) 顶部菜单表格CSS定义(导航菜单) 调用：Class=TopLighNav1
<li>Forum_body(27) 表格边框颜色定义 在下面定义一和二CSS定义控制不到的地方，最好保持和边框CSS定义一中同样的颜色
<li>Forum_body(28) 顶部菜单表格CSS定义(Logo & banner部分) 调用：Class=TopLighNav2


</td>
      <td width="70%" class=tablebody1>
<table cellpadding=6 cellspacing=1 align=center width="100%" style="word-break:break-all;" valign=top>
<% for i=0 to 28%>
<tr ><td class=tableborder1 height=1  ></td></tr>
<tr>
      <td >Forum_body(<%=i%>) 当前值为：<%=Forum_body(i)%>

</td>
    </tr>
<% next %>
  </table>
</td>
    </tr>
    <tr>
      <th height="24" colspan="2"></th>
    </tr>
  </table>

<% end sub 
sub pic()%>

<table cellpadding=6 cellspacing=1 align=center class=tableborder1 style="word-break:break-all;">
    <tr>
      <th height="24" colspan="2">Forum_pic()论坛图片变量</th>
    </tr>
    <tr>
      <td class=tablebody1 colspan="2">
论坛图片变量
</td></tr>
<tr>
    <tr>
      <td width="30%" class=tablebody2 valign=top>
<li>Forum_pic(0) 论坛管理员
<li>Forum_pic(1) 论坛版主
<li>Forum_pic(2) 论坛贵宾
<li>Forum_pic(3) 普通会员
<li>Forum_pic(4) 客人或隐身会员
<li>Forum_pic(5) 突出显示自己的颜色
<li>Forum_pic(6) 常规论坛－－无新帖子
<li>Forum_pic(7) 常规论坛－－有新帖子
<li>Forum_pic(8) 
<li>Forum_pic(9) 
<li>Forum_pic(10) 
<li>Forum_pic(11) 
<li>Forum_pic(12) 
<li>Forum_pic(13) 
<li>Forum_pic(14) 认证论坛－－无新帖子
<li>Forum_pic(15) 认证论坛－－有新帖子


</td>
      <td width="70%" class=tablebody1>
<table cellpadding=6 cellspacing=1 align=center width="100%" style="word-break:break-all;" valign=top>
<% for i=0 to 15%>
<tr ><td class=tableborder1 height=1  ></td></tr>
<tr>
      <td >Forum_pic(<%=i%>) 当前值为：<%=Forum_pic(i)%>    <img src=pic/<%=Forum_pic(i)%> border=0>

</td>
    </tr>
<% next %>
  </table>
</td>
    </tr>
    <tr>
      <th height="24" colspan="2"></th>
    </tr>
  </table>



<% end sub 
sub boardpic()%>
<table cellpadding=6 cellspacing=1 align=center class=tableborder1 style="word-break:break-all;">
    <tr>
      <th height="24" colspan="2">Forum_boardpic()论坛主体图标</th>
    </tr>
    <tr>
      <td class=tablebody1 colspan="2">
<li>
论坛主体图标 

</td></tr>
<tr>
    <tr>
      <td width="30%" class=tablebody2 valign=top>
<li>Forum_boardpic(0) 联盟论坛
<li>Forum_boardpic(1) 发表帖子
<li>Forum_boardpic(2) 发表投票
<li>Forum_boardpic(3) 小字报
<li>Forum_boardpic(4) 回复帖子
<li>Forum_boardpic(5) 帮助
<li>Forum_boardpic(6) 返回首页
<li>Forum_boardpic(7) 返回上级目录
<li>Forum_boardpic(8) 当前目录
<li>Forum_boardpic(9) 新的短消息
<li>Forum_boardpic(10) 我发表的主题
<li>Forum_boardpic(11) 树形浏览帖子模式
<li>Forum_boardpic(12) 平板形浏览帖子模式
<li>Forum_boardpic(13) 下一篇帖子
<li>Forum_boardpic(14) 上一篇帖子
<li>Forum_boardpic(15) 刷新浏览
<li>Forum_boardpic(16) 论坛公告
<hr  size=1>
<li>Forum_statePic(0) 打开的主题
<li>Forum_statePic(1) 热门的主题
<li>Forum_statePic(2) 锁定的主题
<li>Forum_statePic(3) 固顶的主题
<li>Forum_statePic(4) 精华的主题
<li>Forum_statePic(5) 
<li>Forum_statePic(6) 
<li>Forum_statePic(7) 
<li>Forum_statePic(8) 
<li>Forum_statePic(9) 
<li>Forum_statePic(10) 
<li>Forum_statePic(11) 
<li>Forum_statePic(12) 投票的主题
<hr  size=1>
<li>Forum_ubb(0) 粗体
<li>Forum_ubb(1) 斜体
<li>Forum_ubb(2) 下划线
<li>Forum_ubb(3) 居中
<li>Forum_ubb(4) URL连接
<li>Forum_ubb(5) Email地址
<li>Forum_ubb(6) 贴图
<li>Forum_ubb(7) 贴Flash
<li>Forum_ubb(8) 贴ShockWave
<li>Forum_ubb(9) 贴RM影片
<li>Forum_ubb(10) 贴Media Player影片
<li>Forum_ubb(11) 贴QuickTime影片
<li>Forum_ubb(12) 引用文字
<li>Forum_ubb(13) 飞行字
<li>Forum_ubb(14) 移动字
<li>Forum_ubb(15) 发光字
<li>Forum_ubb(16) 阴影字
</td>
      <td width="70%" class=tablebody1>
<table cellpadding=6 cellspacing=1 align=center width="100%" style="word-break:break-all;" valign=top>
<% for i=0 to 16%>
<tr ><td class=tableborder1 height=1  ></td></tr>
<tr>
      <td >Forum_boardpic(<%=i%>) 当前值为：<%=Forum_boardpic(i)%>    <img src=pic/<%=Forum_boardpic(i)%> border=0>

</td>
    </tr>
<% next %>
<% for i=0 to 12%>
<tr ><td class=tableborder1 height=1  ></td></tr>
<tr>
      <td class=tablebody2>Forum_statePic(<%=i%>) 当前值为：<%=Forum_statePic(i)%>    <img src=pic/<%=Forum_statePic(i)%> border=0>

</td>
    </tr>
<% next %>
<% for i=0 to 16%>
<tr ><td class=tableborder1 height=1  ></td></tr>
<tr>
      <td class=tablebody2>Forum_ubb(<%=i%>) 当前值为：<%=Forum_ubb(i)%>    <img src=pic/<%=Forum_ubb(i)%> border=0>

</td>
    </tr>
<% next %>
  </table>
</td>
    </tr>
    <tr>
      <th height="24" colspan="2"></th>
    </tr>
  </table>

<% end sub 
sub topicpic()%>

 <table cellpadding=6 cellspacing=1 align=center class=tableborder1 style="word-break:break-all;">
    <tr>
      <th height="24" colspan="2">Forum_TopicPic()论坛帖子图标 </th>
    </tr>
    <tr>
      <td class=tablebody1 colspan="2">
论坛帖子图标 


</td></tr>
<tr>
    <tr>
      <td width="30%" class=tablebody2 valign=top>
<li>Forum_TopicPic(0) 保存当页为文件
<li>Forum_TopicPic(1) 报告本帖给版主
<li>Forum_TopicPic(2) 显示可打印的版本
<li>Forum_TopicPic(3) 把本帖打包邮递
<li>Forum_TopicPic(4) 把本帖加入论坛收藏
<li>Forum_TopicPic(5) 发送本页面给朋友
<li>Forum_TopicPic(6) 把本帖加入IE收藏
<li>Forum_TopicPic(7) 发送短信给好友
<li>Forum_TopicPic(8) 搜索用户发表帖子
<li>Forum_TopicPic(9) 浏览用户信息
<li>Forum_TopicPic(10) 用户邮箱
<li>Forum_TopicPic(11) 用户OICQ
<li>Forum_TopicPic(12) 用户ICQ
<li>Forum_TopicPic(13) 用户MSN
<li>Forum_TopicPic(14) 用户主页
<li>Forum_TopicPic(15) 引用帖子
<li>Forum_TopicPic(16) 编辑帖子
<li>Forum_TopicPic(17) 删除帖子
<li>Forum_TopicPic(18) 复制帖子
<li>Forum_TopicPic(19) 加入精华帖子
<li>Forum_TopicPic(20) 用户IP
<li>Forum_TopicPic(21) 加为好友
<li>Forum_TopicPic(22) 快速回复

</td>
      <td width="70%" class=tablebody1>
<table cellpadding=6 cellspacing=1 align=center width="100%" style="word-break:break-all;" valign=top>
<% for i=0 to 22%>
<tr ><td class=tableborder1 height=1  ></td></tr>
<tr>
      <td >Forum_TopicPic(<%=i%>) 当前值为：<%=Forum_TopicPic(i)%>    <img src=pic/<%=Forum_TopicPic(i)%> border=0>

</td>
    </tr>
<% next %>
  </table>
</td>
    </tr>
    <tr>
      <th height="24" colspan="2"></th>
    </tr>
  </table>
<% end sub 
sub userset()%>
<table cellpadding=6 cellspacing=1 align=center class=tableborder1 style="word-break:break-all;">
    <tr>
      <th height="24" colspan="2">Forum_user() 用户设定
 </th>
    </tr>
    <tr>
      <td class=tablebody1 colspan="2">
Forum_user() 用户设定
</td></tr>
<tr>
    <tr>
      <td width="30%" class=tablebody2 valign=top>
<li>Forum_user(0) 注册金钱数
<li>Forum_user(1) 发帖增加金钱
<li>Forum_user(2) 跟帖增加金钱
<li>Forum_user(3) 删帖减少金钱
<li>Forum_user(4) 登陆增加金钱
<li>Forum_user(5) 注册经验值
<li>Forum_user(6) 发帖增加经验值
<li>Forum_user(7) 跟帖增加经验值
<li>Forum_user(8) 删帖减少经验值
<li>Forum_user(9) 登陆增加经验值
<li>Forum_user(10) 注册魅力值
<li>Forum_user(11) 发帖增加魅力值
<li>Forum_user(12) 跟帖增加魅力值
<li>Forum_user(13) 删帖减少魅力值
<li>Forum_user(14) 登陆增加魅力值
<li>Forum_user(15) 精华增加金钱
<li>Forum_user(16) 精华增加魅力值
<li>Forum_user(17) 精华增加经验值
</td>
      <td width="70%" class=tablebody1>
<table cellpadding=6 cellspacing=1 align=center width="100%" style="word-break:break-all;" valign=top>
<% for i=0 to 17%>
<tr ><td class=tableborder1 height=1  ></td></tr>
<tr>
      <td >Forum_user(<%=i%>) 当前值为：<%=Forum_user(i)%> 

</td>
    </tr>
<% next %>
  </table>
</td>
    </tr>
    <tr>
      <th height="24" colspan="2"></th>
    </tr>
  </table>

<%end sub


sub boardset()%>
<table cellpadding=6 cellspacing=1 align=center class=tableborder1 style="word-break:break-all;">
    <tr>
      <th height="24" colspan="2">Board_Setting() 分版论坛设定
 </th>
    </tr>
    <tr>
      <td class=tablebody1 colspan="2">
Board_Setting() 分版论坛设定<br>
0,0,0,0,1,0,1,1,1,1,1,1,1,1,1,1,16240,3,300,gif|jpg|jpeg|bmp|png|rar|txt|zip,,0,0|24,1,1,300,20,10,12,15,1,20,10,0,0,0,0,0,0,1,0,4,0,0,0,0,0,0,0,0,0
</td></tr>
<tr>
    <tr>
      <td width="70%" class=tablebody2 valign=top>
<li>Board_Setting(0)、论坛开放/锁定
<li>Board_Setting(1)、论坛隐含/不隐含（用户组中设定是否可进入隐含论坛）
<li>Board_Setting(2)、论坛加密/不加密（加密论坛需要设定可进入用户）
<li>Board_Setting(3)、论坛审核不开放/开放
<li>Board_Setting(4)、发贴模式简单/高级
<li>Board_Setting(5)、HTML支持开启/关闭
<li>Board_Setting(6)、UBB功能开启/关闭
<li>Board_Setting(7)、贴图标签开启/关闭
<li>Board_Setting(8)、表情标签开启/关闭
<li>Board_Setting(9)、多媒体标签开启/关闭（包括Flash,RM,AVI等）
<li>Board_Setting(10)、金钱贴开启/关闭
<li>Board_Setting(11)、积分
<li>Board_Setting(12)、魅力
<li>Board_Setting(13)、威望
<li>Board_Setting(14)、文章
<li>Board_Setting(15)、回复可见
<li>Board_Setting(16)、贴子内容最大字节数
<li>Board_Setting(17)、发贴后返回（首页、论坛、贴子）
<li>Board_Setting(18)、允许同时在线数
<li>Board_Setting(19)、上传文件类型
<li>Board_Setting(20)、上传文件目录＝＝取消
<li>Board_Setting(21)、启动版面定时开放功能
<li>Board_Setting(22)、论坛开放时间（0－24点）
<li>Board_Setting(23)、版面语言（简体中文、繁体中文、英文）
<li>Board_Setting(24)、论坛风格（讨论区风格、默认风格）
<li>Board_Setting(25)、讨论区风格预览字数
<li>Board_Setting(26)、贴子列表每页帖子数
<li>Board_Setting(27)、浏览贴子每页帖子数
<li>Board_Setting(28)、帖子正文字号
<li>Board_Setting(29)、帖子正文行间距
<li>Board_Setting(30)、防灌水机制
<li>Board_Setting(31)、每次发贴时间间隔
<li>Board_Setting(32)、最多投票项目
<li>Board_Setting(33)、主版主可以增删副版主
<li>Board_Setting(34)、主版主可以修改CSS设置
<li>Board_Setting(35)、所有版主可以修改CSS设置
<li>Board_Setting(36)、是否对论坛事件中的操作者保密（对管理员无效）
<li>Board_Setting(37)、论坛默认读取帖子量（按所需时间内）
<li>Board_Setting(38)、默认排序（最后回复时间、标题、发贴人、回复数、浏览数、发表时间）
<li>Board_Setting(39)、备用字段1（是否采用简洁风格列表）
<li>Board_Setting(40)、备用字段2（是否继承上级版主权限）
<li>Board_Setting(41)、备用字段3（简洁风格列表一行版面数）
<li>Board_Setting(42)、备用字段4（是否显示下级论坛帖子）
<li>Board_Setting(43)、备用字段5（作为分类论坛不允许发贴）
<li>Board_Setting(44)、备用字段6
<li>Board_Setting(45)、备用字段7
<li>Board_Setting(46)、备用字段8
<li>Board_Setting(47)、备用字段9
<li>Board_Setting(48)、备用字段10
<li>Board_Setting(49)、备用字段11
<li>Board_Setting(50)、备用字段12

</td>
      <td width="30%" class=tablebody1>
<table cellpadding=6 cellspacing=1 align=center width="100%" style="word-break:break-all;" valign=top>
<% for i=0 to 50%>
<tr ><td class=tableborder1 height=1  ></td></tr>
<tr>
      <td >Board_Setting(<%=i%>) 当前值为：[<%=Board_Setting(i)%>]

</td>
    </tr>
<% next %>
  </table>
</td>
    </tr>
    <tr>
      <th height="24" colspan="2"></th>
    </tr>
  </table>

<%end sub%>