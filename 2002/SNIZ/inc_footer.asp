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
    </div> 
    <td>
    <table width=100% border=0  cellpadding="0" cellspacing = "4"> 
      <tr >
        <td bgcolor="">
        <table border=0 width="100%" align="center" cellpadding="1" cellspacing="0">
          <tr>
            <td bgcolor="<% =strForumCellColor %>" align=center valign=top nowrap><% =strCopyright %></td>
			<td bgcolor="<% =strForumCellColor %>" width=10 nowrap><a href="#top"><img src="<%=strImageURL %>icon_go_up.gif" height=15 width=15 border="0" align="right" alt="����"></a></font></td>    
          </tr>
        </table>
        </td>
      </tr>
    </table>
    </td>
  </tr>
  <tr>
	<td>
	  <table border=0 width="100%" align="center" cellpadding="0" cellspacing="0">
	  <tr>
        <td align="left">
        <acronym title="<% =strVersion %> �����޸�:��Դ����վ"><% if strShowImagePoweredBy = "1" then %><% else %><a href="http://jarod.y365.com" target=_blank>����αװ�ռ��޸�֧��</a><% end if %></acronym>
        <td align="right">
        <acronym title="<% =strVersion %>�������İ�"><% if strShowImagePoweredBy = "1" then %><a href="http://jarod.y365.com" target=_blank>����αװ�ռ��޸�֧��</a><% else %><Script Src="http://down.my0754.com/logo.js"></Script><% end if %></acronym>
        </td>
      </tr>
	</table>
	</td>
  </tr>
</table>



</font>

</BODY>
</HTML>
<% my_Conn.Close %>
<% set my_Conn = nothing %>