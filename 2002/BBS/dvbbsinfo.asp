<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->

<%
boardid=request("boardid")
if not isnumeric(boardid) or boardid="" or boardid=0 then
boardid=1
end if


Stats="��̳������Ϣ"
call activeonline()
	call nav()
	call head_var(1,BoardDepth,0,0)


%>

<p>
</p>
<table cellpadding=6 cellspacing=1 align=center class=tableborder2>
    <tr>
      <td class=tablebody1>��̳������Ϣ</td>
    </tr>
    <tr>
      <td class=tablebody1>
<a href=?view=Setting&boardid=<%=boardid%> >����������Ϣ</a> | 
<a href=?view=info&boardid=<%=boardid%> >������Ϣ</a> | 
<a href=?view=Group&boardid=<%=boardid%> >�û���Ȩ������</a> | 
<a href=?view=css&boardid=<%=boardid%> >������ɫCSS�ı���</a> | 
<a href=?view=pic&boardid=<%=boardid%> >��̳ͼƬ����</a> | 
<a href=?view=boardpic&boardid=<%=boardid%> >��̳����ͼ��</a> | 
<a href=?view=TopicPic&boardid=<%=boardid%> >��̳����ͼ��</a> | 
<a href=?view=userset&boardid=<%=boardid%> >�û��趨</a> | 
<a href=?view=boardset&boardid=<%=boardid%> >�ְ��趨</a>
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
      <th height="24" colspan="2">Forum_Setting()  ����������Ϣ</th>
    </tr>
    <tr>
      <td class=tablebody1 colspan="2">
Forum_Setting()  ����������Ϣ  ��ǰֵ��������<br>
0,300,1,1,0,0,1,1,20,0,0,20,10,16240,0,0,0,1,0,0,30,0,3000,0,1,0,500,0,1,0,1,1,1,100,1,0,0,1,0,10,0,0,1,10,0,0,1,1,1,0,0,0,0,0,0,1,1,0,1000,9,15,4,0,300,0,1,0,1,20,3,1,1,1
</td></tr>
<tr>
    <tr>
      <td width="70%" class=tablebody2 valign=top>
<li>forum_setting(0) ������ʱ��    -ok
<li>forum_setting(1) �ű���ʱʱ��    -ok
<li>forum_setting(2) �����ʼ����  0����֧�� 1��JMAIL 2��CDONTS  3��ASPEMAIL-ok
<li>forum_setting(3) �Ƿ����ü�¼�����Ķ��û�����
<li>forum_setting(4) ��ҳģʽ   0-�°���ʽ 1-�ɰ���ʽ  
<li>forum_setting(5) ��ҳ��ʾ��̳���-ok
<li>forum_setting(6) �û�ͷ��   0���ر� 1����  
<li>forum_setting(7) ͷ���ϴ�   0���ر� 1����  
<li>forum_setting(8) ɾ������û�ʱ�� -ok   
<li>forum_setting(9) 
<li>forum_setting(10) �¶���Ϣ��������  0���� 1����  
<li>forum_setting(11) ÿҳ��ʾ����¼    ==һ���ҳ
<li>forum_setting(12) 
<li>forum_setting(13)
<li>forum_setting(14) ����������ʾ�û�����  0���ر� 1����  -ok
<li>forum_setting(15) ����������ʾ��������  0���ر� 1����  -ok
<li>forum_setting(16)    
<li>forum_setting(17) 
<li>forum_setting(18)    
<li>forum_setting(19) ��ˢ�»���   0���ر� 1���� -ok
<li>forum_setting(20) ���ˢ��ʱ����   -ok
<li>forum_setting(21) ��̳��ǰ״̬   -ok
<li>forum_setting(22) ͬһIPע����ʱ��-ok   
<li>forum_setting(23) Email֪ͨ����   0���ر� 1���� 
<li>forum_setting(24) һ��Emailֻ��ע��һ���ʺ� 0���ر� 1���� 
<li>forum_setting(25) ע����Ҫ����Ա��֤  0���ر� 1���� 
<li>forum_setting(26) ����ͬʱ������   -ok
<li>forum_setting(27) 
<li>forum_setting(28) �����б��Ƿ���ʾ�û�IP
<li>forum_setting(29) �Ƿ���ʾ�����ջ�Ա  0���� 1���� 
<li>forum_setting(30) �Ƿ���ʾҳ��ִ��ʱ��  0���� 1���� 
<li>forum_setting(31) ���ٵ�½λ��   1������ 2���ײ� ��0��������ʾ
<li>forum_setting(32) �Ƿ�����̳����  0���� 1����
<li>forum_setting(33) ���������б���ʾ��ǰλ��
<li>forum_setting(34) ���������б���ʾ��½�ͻʱ��
<li>forum_setting(35) ���������б���ʾ������Ͳ���ϵͳ
<li>forum_setting(36) ���������б���ʾ��Դ���رպ��������û������û��������˿��Բ鿴������û��ɲ鿴��Դ��
<li>forum_setting(37) �Ƿ��������û�ע��
<li>forum_setting(38) 
<li>forum_setting(39) 
<li>forum_setting(40) ����û�������
<li>forum_setting(41) ��û�������  
<li>forum_setting(42) �������ǩ��
<li>forum_setting(43) 
<li>forum_setting(44) ��Ϊ���Ż�����������ֵ  
<li>forum_setting(45)   
<li>forum_setting(46) �������Ż�ӭ��ע���û�  0���ر� 1����
<li>forum_setting(47) ����ע����Ϣ�ʼ�  0���ر� 1����
<li>forum_setting(48) �༭����������ʾ"��xxx��yyy�༭"����Ϣ
<li>forum_setting(49) ����Ա�༭����ʾ"��XXX�༭"����Ϣ
<li>forum_setting(50) �ȴ�"��XXX�༭"��Ϣ��ʾ��ʱ��
<li>forum_setting(51) �༭����ʱ��   
<li>forum_setting(52) ��ֹ���ʼ���ַ
<li>forum_setting(53) �����û�ʹ��ͷ��    --����������趨
<li>forum_setting(54) ʹ���Զ���ͷ������ٷ�����    
<li>forum_setting(55) ���������վ���ϴ�ͷ��
<li>forum_setting(56) ��������ͷ���ļ���С
<li>forum_setting(57) ���ͷ��ߴ�
<li>forum_setting(58) �鿴ͼƬѡ��
<li>forum_setting(59) �û�ͷ����󳤶�
<li>forum_setting(60) �Զ���ͷ�����ٷ�����������
<li>forum_setting(61) �Զ���ͷ��ע����������
<li>forum_setting(62) �Զ���ͷ������������������һ������
<li>forum_setting(63) �Զ���ͷ����Ҫ���εĴ���
<li>forum_setting(64) ��ˢ�¹�����Ч��ҳ��
<li>forum_setting(65) �û�ǩ���Ƿ���UBB����  0���ر� 1���� 
<li>forum_setting(66) �û�ǩ���Ƿ���HTML���� 0���ر� 1���� 
<li>forum_setting(67) �û��Ƿ�����ͼ��ǩ  0���ر� 1���� 
<li>forum_setting(68) �û������б����   
<li>forum_setting(69) �Ƿ�ʱ������̳
<li>forum_setting(70) ��ʱ������̳����
<li>forum_setting(71) 
<li>forum_setting(72) 
</td>
      <td width="30%" class=tablebody1>
<table cellpadding=6 cellspacing=1 align=center width="100%" style="word-break:break-all;" valign=top>
<% for i=0 to 72%>
<tr ><td class=tableborder1 height=1  ></td></tr>
<tr>
      <td >Forum_Setting(<%=i%>) ��ǰֵΪ��<%=Forum_Setting(i)%>

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
      <th height="24" colspan="2">forum_info()  ������Ϣ</th>
    </tr>
    <tr>
      <td class=tablebody1 colspan="2">
forum_info()   ������Ϣ  Ĭ����Ϣ����
<li>
�����ȷ���̳,http://www.dvbbs.net/,�����ȷ�,http://www.aspsky.net/,61.145.114.64,club@aspsky.net,images/LOGO.GIF,pic/,face/,����ʱ��


</td></tr>
<tr>
    <tr>
      <td width="30%" class=tablebody2 valign=top>
<li>Forum_info(0) ��̳����
<li>Forum_info(1) ��̳��url
<li>Forum_info(2) ��ҳ����
<li>Forum_info(3) ��ҳURL
<li>Forum_info(4) SMTP Server��ַ
<li>Forum_info(5) ��̳����ԱEmail
<li>Forum_info(6) ��̳��ҳLogo��ַ
<li>Forum_info(7) ��̳ͼƬĿ¼
<li>Forum_info(8) ��̳����Ŀ¼
<li>Forum_info(9) ��̳����ʱ��
<li>Forum_info(10) ��������Ŀ¼
<li>Forum_info(11) ��̳ͷ��Ŀ¼



</td>
      <td width="70%" class=tablebody1>
<table cellpadding=6 cellspacing=1 align=center width="100%" style="word-break:break-all;" valign=top>
<% for i=0 to 11%>
<tr ><td class=tableborder1 height=1  ></td></tr>
<tr>
      <td >Forum_info(<%=i%>) ��ǰֵΪ��<%=Forum_info(i)%>

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
      <th height="24" colspan="2">GroupSetting()[reGroupSetting()]  �û���Ȩ������</th>
    </tr>
    <tr>
      <td width="75%" class=tablebody2 valign=top>
GroupSetting()[reGroupSetting()]  �û���Ȩ������
<li>
<li>UserGroupID title
<li>1  ����Ա
<li>2  ��������
<li>3  ����
<li>4  ע���û�
<li>5  �ȴ���֤��(COPPA)��Ա
<li>6  �ȴ��ʼ���֤�Ļ�Ա
<li>7  δע��/δ��¼�û�
<li>8  ���
<hr class=tableborder1>
<li>GroupSetting(0)  ���������̳   0���� 1����
<li>GroupSetting(1)  ���Բ鿴��Ա��Ϣ  0���� 1����
<li>GroupSetting(2)  ���Բ鿴�����˷��������� 0���� 1����
<li>GroupSetting(3)  ���Է���������   0���� 1����
<li>GroupSetting(4)  ���Իظ��Լ�������  0���� 1����
<li>GroupSetting(5)  ���Իظ������˵�����  0���� 1����
<li>GroupSetting(6)  ��������̳�������ֵ�ʱ���������0���� 1����
<li>GroupSetting(7)  �����ϴ�����   0���� 1����
<li>GroupSetting(8)  ���Է�����ͶƱ   0���� 1����
<li>GroupSetting(9)  ���Բ���ͶƱ   0���� 1����
<li>GroupSetting(10) ���Ա༭�Լ�������  0���� 1����
<li>GroupSetting(11) ����ɾ���Լ�������  0���� 1����
<li>GroupSetting(12) �����ƶ��Լ������ӵ�������̳ 0���� 1����
<li>GroupSetting(13) �Դ�/�ر��Լ����������� 0���� 1����
<li>GroupSetting(14) ����������̳   0���� 1����
<li>GroupSetting(15) ����ʹ��''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''���ͱ�ҳ������''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''���� 0���� 1����
<li>GroupSetting(16) �����޸ĸ�������  0���� 1����
<li>GroupSetting(17) ���Է���С�ֱ�   0���� 1����
<li>GroupSetting(18) ����ɾ������������  0���� 1����
<li>GroupSetting(19) �����ƶ�����������  0���� 1����
<li>GroupSetting(20) ���Դ�/�ر�����������  0���� 1����
<li>GroupSetting(21) ���Թ̶�/����̶�����  0���� 1����
<li>GroupSetting(22) ���Խ���/�ͷ������û�  0���� 1����
<li>GroupSetting(23) ���Ա༭����������  0���� 1����
<li>GroupSetting(24) ���Լ���/�����������  0���� 1����
<li>GroupSetting(25) ���Է�������   0���� 1����
<li>GroupSetting(26) ���Թ�����   0���� 1����
<li>GroupSetting(27) ���Թ���С�ֱ�   0���� 1����
<li>GroupSetting(28) ��������/����/��������û� 0���� 1����
<li>GroupSetting(29) ����ɾ���û�1��10������������ 0���� 1����
<li>GroupSetting(30) ���Բ鿴����IP����Դ  0���� 1����
<li>GroupSetting(31) �����޶�IP����   0���� 1����
<li>GroupSetting(32) ���Է��Ͷ���   0���� 1����
<li>GroupSetting(33) ��෢���û�  
<li>GroupSetting(34) �������ݴ�С����  
<li>GroupSetting(35) �����С����  
<li>GroupSetting(36) �Ƿ���������ӵ�Ȩ�� 0���� 1����
<li>GroupSetting(37) �Ƿ��н���������̳��Ȩ�� 0���� 1����
<li>GroupSetting(38) �����п��Կ��Բ鿴������ǩ�� 0���� 1����
<li>GroupSetting(39) ���������̳�¼�  0���� 1����
<li>GroupSetting(40) ����ϴ��ļ�����  
<li>GroupSetting(41) ���������������  0���� 1����
<li>GroupSetting(42) ���Թ����û�Ȩ��  0���� 1����
<li>GroupSetting(43) ���Խ���/�ͷ��û�  0���� 1����
<li>GroupSetting(44) �ϴ��ļ���С����  
<li>GroupSetting(45) ��������ɾ�����ӣ�ǰ̨�� 0���� 1����
<li>GroupSetting(46) ����С�ֱ������Ǯ  
<li>GroupSetting(47) �������������Ǯ


</td>
      <td width="25%" class=tablebody1>
<table cellpadding=6 cellspacing=1 align=center width="100%" style="word-break:break-all;" valign=top>
<% for i=0 to 47%>
<tr ><td class=tableborder1 height=1  ></td></tr>
<tr>
      <td >GroupSetting(<%=i%>) ��ǰֵΪ��<%=GroupSetting(i)%>

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
      <th height="24" colspan="2">Forum_body()  ������ɫCSS�ı���
</th>
    </tr>
    <tr>
      <td class=tablebody1 colspan="2">
Forum_body()  ������ɫCSS�ı���

</td></tr>
<tr>
    <tr>
      <td width="30%" class=tablebody2 valign=top>
<li>Forum_body()  ������ɫCSS�ı���
<li>Forum_body(0) ���߿��壩CSS����һ ���ã�Class=TableBorder1
<li>Forum_body(1) ���߿��壩CSS����� ���ã�Class=TableBorder2
<li>Forum_body(2) ������CSS����һ������� ���ã�th
<li>Forum_body(3) ������CSS�������ǳ������ ���ã�Class=TableTitle2
<li>Forum_body(4) �����CSS����һ ���ã�Class=TableBody1
<li>Forum_body(5) �������ɫ��(1��2��ɫ����ʾ�д���) ���ã�Class=TableBody2
<li>Forum_body(6) ��̳�����ܵ�CSS���� 
<li>Forum_body(7) ��������������CSS���壨����� ���ã�id=TableTitleLink
<li>Forum_body(8) ��������������ɫ 
<li>Forum_body(9) ��ʾ���ӵ�ʱ��������ӣ�ת�����ӣ��ظ��ȵ���ɫ 
<li>Forum_body(10) ��̳����ܵ�CSS���� 
<li>Forum_body(11) ��̳BODY��ǩ 
<li>Forum_body(12) ����� 
<li>Forum_body(13)  
<li>Forum_body(14) ��ҳ������ɫ 
<li>Forum_body(15) һ���û�����������ɫ 
<li>Forum_body(16) һ���û������ϵĹ�����ɫ 
<li>Forum_body(17) ��������������ɫ 
<li>Forum_body(18) ���������ϵĹ�����ɫ 
<li>Forum_body(19) ����Ա����������ɫ 
<li>Forum_body(20) ���������ϵĹ�����ɫ 
<li>Forum_body(21) �������������ɫ 
<li>Forum_body(22) ��������ϵĹ�����ɫ 
<li>Forum_body(23) �����˵����CSS����(Logo & Banner�·�) ���ã�Class=TopLighNav
<li>Forum_body(24) �����˵����CSS����(Logo & Banner�Ϸ�) ���ã�Class=TopDarkNav
<li>Forum_body(25) Body��CSS���� 
<li>Forum_body(26) �����˵����CSS����(�����˵�) ���ã�Class=TopLighNav1
<li>Forum_body(27) ���߿���ɫ���� �����涨��һ�Ͷ�CSS������Ʋ����ĵط�����ñ��ֺͱ߿�CSS����һ��ͬ������ɫ
<li>Forum_body(28) �����˵����CSS����(Logo & banner����) ���ã�Class=TopLighNav2


</td>
      <td width="70%" class=tablebody1>
<table cellpadding=6 cellspacing=1 align=center width="100%" style="word-break:break-all;" valign=top>
<% for i=0 to 28%>
<tr ><td class=tableborder1 height=1  ></td></tr>
<tr>
      <td >Forum_body(<%=i%>) ��ǰֵΪ��<%=Forum_body(i)%>

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
      <th height="24" colspan="2">Forum_pic()��̳ͼƬ����</th>
    </tr>
    <tr>
      <td class=tablebody1 colspan="2">
��̳ͼƬ����
</td></tr>
<tr>
    <tr>
      <td width="30%" class=tablebody2 valign=top>
<li>Forum_pic(0) ��̳����Ա
<li>Forum_pic(1) ��̳����
<li>Forum_pic(2) ��̳���
<li>Forum_pic(3) ��ͨ��Ա
<li>Forum_pic(4) ���˻������Ա
<li>Forum_pic(5) ͻ����ʾ�Լ�����ɫ
<li>Forum_pic(6) ������̳������������
<li>Forum_pic(7) ������̳������������
<li>Forum_pic(8) 
<li>Forum_pic(9) 
<li>Forum_pic(10) 
<li>Forum_pic(11) 
<li>Forum_pic(12) 
<li>Forum_pic(13) 
<li>Forum_pic(14) ��֤��̳������������
<li>Forum_pic(15) ��֤��̳������������


</td>
      <td width="70%" class=tablebody1>
<table cellpadding=6 cellspacing=1 align=center width="100%" style="word-break:break-all;" valign=top>
<% for i=0 to 15%>
<tr ><td class=tableborder1 height=1  ></td></tr>
<tr>
      <td >Forum_pic(<%=i%>) ��ǰֵΪ��<%=Forum_pic(i)%>    <img src=pic/<%=Forum_pic(i)%> border=0>

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
      <th height="24" colspan="2">Forum_boardpic()��̳����ͼ��</th>
    </tr>
    <tr>
      <td class=tablebody1 colspan="2">
<li>
��̳����ͼ�� 

</td></tr>
<tr>
    <tr>
      <td width="30%" class=tablebody2 valign=top>
<li>Forum_boardpic(0) ������̳
<li>Forum_boardpic(1) ��������
<li>Forum_boardpic(2) ����ͶƱ
<li>Forum_boardpic(3) С�ֱ�
<li>Forum_boardpic(4) �ظ�����
<li>Forum_boardpic(5) ����
<li>Forum_boardpic(6) ������ҳ
<li>Forum_boardpic(7) �����ϼ�Ŀ¼
<li>Forum_boardpic(8) ��ǰĿ¼
<li>Forum_boardpic(9) �µĶ���Ϣ
<li>Forum_boardpic(10) �ҷ��������
<li>Forum_boardpic(11) �����������ģʽ
<li>Forum_boardpic(12) ƽ�����������ģʽ
<li>Forum_boardpic(13) ��һƪ����
<li>Forum_boardpic(14) ��һƪ����
<li>Forum_boardpic(15) ˢ�����
<li>Forum_boardpic(16) ��̳����
<hr  size=1>
<li>Forum_statePic(0) �򿪵�����
<li>Forum_statePic(1) ���ŵ�����
<li>Forum_statePic(2) ����������
<li>Forum_statePic(3) �̶�������
<li>Forum_statePic(4) ����������
<li>Forum_statePic(5) 
<li>Forum_statePic(6) 
<li>Forum_statePic(7) 
<li>Forum_statePic(8) 
<li>Forum_statePic(9) 
<li>Forum_statePic(10) 
<li>Forum_statePic(11) 
<li>Forum_statePic(12) ͶƱ������
<hr  size=1>
<li>Forum_ubb(0) ����
<li>Forum_ubb(1) б��
<li>Forum_ubb(2) �»���
<li>Forum_ubb(3) ����
<li>Forum_ubb(4) URL����
<li>Forum_ubb(5) Email��ַ
<li>Forum_ubb(6) ��ͼ
<li>Forum_ubb(7) ��Flash
<li>Forum_ubb(8) ��ShockWave
<li>Forum_ubb(9) ��RMӰƬ
<li>Forum_ubb(10) ��Media PlayerӰƬ
<li>Forum_ubb(11) ��QuickTimeӰƬ
<li>Forum_ubb(12) ��������
<li>Forum_ubb(13) ������
<li>Forum_ubb(14) �ƶ���
<li>Forum_ubb(15) ������
<li>Forum_ubb(16) ��Ӱ��
</td>
      <td width="70%" class=tablebody1>
<table cellpadding=6 cellspacing=1 align=center width="100%" style="word-break:break-all;" valign=top>
<% for i=0 to 16%>
<tr ><td class=tableborder1 height=1  ></td></tr>
<tr>
      <td >Forum_boardpic(<%=i%>) ��ǰֵΪ��<%=Forum_boardpic(i)%>    <img src=pic/<%=Forum_boardpic(i)%> border=0>

</td>
    </tr>
<% next %>
<% for i=0 to 12%>
<tr ><td class=tableborder1 height=1  ></td></tr>
<tr>
      <td class=tablebody2>Forum_statePic(<%=i%>) ��ǰֵΪ��<%=Forum_statePic(i)%>    <img src=pic/<%=Forum_statePic(i)%> border=0>

</td>
    </tr>
<% next %>
<% for i=0 to 16%>
<tr ><td class=tableborder1 height=1  ></td></tr>
<tr>
      <td class=tablebody2>Forum_ubb(<%=i%>) ��ǰֵΪ��<%=Forum_ubb(i)%>    <img src=pic/<%=Forum_ubb(i)%> border=0>

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
      <th height="24" colspan="2">Forum_TopicPic()��̳����ͼ�� </th>
    </tr>
    <tr>
      <td class=tablebody1 colspan="2">
��̳����ͼ�� 


</td></tr>
<tr>
    <tr>
      <td width="30%" class=tablebody2 valign=top>
<li>Forum_TopicPic(0) ���浱ҳΪ�ļ�
<li>Forum_TopicPic(1) ���汾��������
<li>Forum_TopicPic(2) ��ʾ�ɴ�ӡ�İ汾
<li>Forum_TopicPic(3) �ѱ�������ʵ�
<li>Forum_TopicPic(4) �ѱ���������̳�ղ�
<li>Forum_TopicPic(5) ���ͱ�ҳ�������
<li>Forum_TopicPic(6) �ѱ�������IE�ղ�
<li>Forum_TopicPic(7) ���Ͷ��Ÿ�����
<li>Forum_TopicPic(8) �����û���������
<li>Forum_TopicPic(9) ����û���Ϣ
<li>Forum_TopicPic(10) �û�����
<li>Forum_TopicPic(11) �û�OICQ
<li>Forum_TopicPic(12) �û�ICQ
<li>Forum_TopicPic(13) �û�MSN
<li>Forum_TopicPic(14) �û���ҳ
<li>Forum_TopicPic(15) ��������
<li>Forum_TopicPic(16) �༭����
<li>Forum_TopicPic(17) ɾ������
<li>Forum_TopicPic(18) ��������
<li>Forum_TopicPic(19) ���뾫������
<li>Forum_TopicPic(20) �û�IP
<li>Forum_TopicPic(21) ��Ϊ����
<li>Forum_TopicPic(22) ���ٻظ�

</td>
      <td width="70%" class=tablebody1>
<table cellpadding=6 cellspacing=1 align=center width="100%" style="word-break:break-all;" valign=top>
<% for i=0 to 22%>
<tr ><td class=tableborder1 height=1  ></td></tr>
<tr>
      <td >Forum_TopicPic(<%=i%>) ��ǰֵΪ��<%=Forum_TopicPic(i)%>    <img src=pic/<%=Forum_TopicPic(i)%> border=0>

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
      <th height="24" colspan="2">Forum_user() �û��趨
 </th>
    </tr>
    <tr>
      <td class=tablebody1 colspan="2">
Forum_user() �û��趨
</td></tr>
<tr>
    <tr>
      <td width="30%" class=tablebody2 valign=top>
<li>Forum_user(0) ע���Ǯ��
<li>Forum_user(1) �������ӽ�Ǯ
<li>Forum_user(2) �������ӽ�Ǯ
<li>Forum_user(3) ɾ�����ٽ�Ǯ
<li>Forum_user(4) ��½���ӽ�Ǯ
<li>Forum_user(5) ע�ᾭ��ֵ
<li>Forum_user(6) �������Ӿ���ֵ
<li>Forum_user(7) �������Ӿ���ֵ
<li>Forum_user(8) ɾ�����پ���ֵ
<li>Forum_user(9) ��½���Ӿ���ֵ
<li>Forum_user(10) ע������ֵ
<li>Forum_user(11) ������������ֵ
<li>Forum_user(12) ������������ֵ
<li>Forum_user(13) ɾ����������ֵ
<li>Forum_user(14) ��½��������ֵ
<li>Forum_user(15) �������ӽ�Ǯ
<li>Forum_user(16) ������������ֵ
<li>Forum_user(17) �������Ӿ���ֵ
</td>
      <td width="70%" class=tablebody1>
<table cellpadding=6 cellspacing=1 align=center width="100%" style="word-break:break-all;" valign=top>
<% for i=0 to 17%>
<tr ><td class=tableborder1 height=1  ></td></tr>
<tr>
      <td >Forum_user(<%=i%>) ��ǰֵΪ��<%=Forum_user(i)%> 

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
      <th height="24" colspan="2">Board_Setting() �ְ���̳�趨
 </th>
    </tr>
    <tr>
      <td class=tablebody1 colspan="2">
Board_Setting() �ְ���̳�趨<br>
0,0,0,0,1,0,1,1,1,1,1,1,1,1,1,1,16240,3,300,gif|jpg|jpeg|bmp|png|rar|txt|zip,,0,0|24,1,1,300,20,10,12,15,1,20,10,0,0,0,0,0,0,1,0,4,0,0,0,0,0,0,0,0,0
</td></tr>
<tr>
    <tr>
      <td width="70%" class=tablebody2 valign=top>
<li>Board_Setting(0)����̳����/����
<li>Board_Setting(1)����̳����/���������û������趨�Ƿ�ɽ���������̳��
<li>Board_Setting(2)����̳����/�����ܣ�������̳��Ҫ�趨�ɽ����û���
<li>Board_Setting(3)����̳��˲�����/����
<li>Board_Setting(4)������ģʽ��/�߼�
<li>Board_Setting(5)��HTML֧�ֿ���/�ر�
<li>Board_Setting(6)��UBB���ܿ���/�ر�
<li>Board_Setting(7)����ͼ��ǩ����/�ر�
<li>Board_Setting(8)�������ǩ����/�ر�
<li>Board_Setting(9)����ý���ǩ����/�رգ�����Flash,RM,AVI�ȣ�
<li>Board_Setting(10)����Ǯ������/�ر�
<li>Board_Setting(11)������
<li>Board_Setting(12)������
<li>Board_Setting(13)������
<li>Board_Setting(14)������
<li>Board_Setting(15)���ظ��ɼ�
<li>Board_Setting(16)��������������ֽ���
<li>Board_Setting(17)�������󷵻أ���ҳ����̳�����ӣ�
<li>Board_Setting(18)������ͬʱ������
<li>Board_Setting(19)���ϴ��ļ�����
<li>Board_Setting(20)���ϴ��ļ�Ŀ¼����ȡ��
<li>Board_Setting(21)���������涨ʱ���Ź���
<li>Board_Setting(22)����̳����ʱ�䣨0��24�㣩
<li>Board_Setting(23)���������ԣ��������ġ��������ġ�Ӣ�ģ�
<li>Board_Setting(24)����̳������������Ĭ�Ϸ��
<li>Board_Setting(25)�����������Ԥ������
<li>Board_Setting(26)�������б�ÿҳ������
<li>Board_Setting(27)���������ÿҳ������
<li>Board_Setting(28)�����������ֺ�
<li>Board_Setting(29)�����������м��
<li>Board_Setting(30)������ˮ����
<li>Board_Setting(31)��ÿ�η���ʱ����
<li>Board_Setting(32)�����ͶƱ��Ŀ
<li>Board_Setting(33)��������������ɾ������
<li>Board_Setting(34)�������������޸�CSS����
<li>Board_Setting(35)�����а��������޸�CSS����
<li>Board_Setting(36)���Ƿ����̳�¼��еĲ����߱��ܣ��Թ���Ա��Ч��
<li>Board_Setting(37)����̳Ĭ�϶�ȡ��������������ʱ���ڣ�
<li>Board_Setting(38)��Ĭ���������ظ�ʱ�䡢���⡢�����ˡ��ظ����������������ʱ�䣩
<li>Board_Setting(39)�������ֶ�1���Ƿ���ü�����б�
<li>Board_Setting(40)�������ֶ�2���Ƿ�̳��ϼ�����Ȩ�ޣ�
<li>Board_Setting(41)�������ֶ�3��������б�һ�а�������
<li>Board_Setting(42)�������ֶ�4���Ƿ���ʾ�¼���̳���ӣ�
<li>Board_Setting(43)�������ֶ�5����Ϊ������̳����������
<li>Board_Setting(44)�������ֶ�6
<li>Board_Setting(45)�������ֶ�7
<li>Board_Setting(46)�������ֶ�8
<li>Board_Setting(47)�������ֶ�9
<li>Board_Setting(48)�������ֶ�10
<li>Board_Setting(49)�������ֶ�11
<li>Board_Setting(50)�������ֶ�12

</td>
      <td width="30%" class=tablebody1>
<table cellpadding=6 cellspacing=1 align=center width="100%" style="word-break:break-all;" valign=top>
<% for i=0 to 50%>
<tr ><td class=tableborder1 height=1  ></td></tr>
<tr>
      <td >Board_Setting(<%=i%>) ��ǰֵΪ��[<%=Board_Setting(i)%>]

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