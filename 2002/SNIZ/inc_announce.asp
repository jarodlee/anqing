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

'## Forum_SQL - Build SQL to get forums via category
strSql = "SELECT " & strTablePrefix & "ANNOUNCE.A_ID, "
strSql = strSql & strTablePrefix & "ANNOUNCE.A_AUTHOR, "
strSql = strSql & strTablePrefix & "ANNOUNCE.A_SUBJECT, "
strSql = strSql & strTablePrefix & "ANNOUNCE.A_MESSAGE, "
strSql = strSql & strTablePrefix & "ANNOUNCE.A_START_DATE, "
strSql = strSql & strTablePrefix & "ANNOUNCE.A_END_DATE "
strSql = strSql & "FROM " & strTablePrefix & "ANNOUNCE "
strSql = strSql & " WHERE " & strTablePrefix & "ANNOUNCE.A_START_DATE <= " & "'" & DatetoStr(Now()) & "'"
strSql = strSql & " AND " & strTablePrefix & "ANNOUNCE.A_END_DATE > " & "'" & DatetoStr(Now()) & "'"
strSql = strSql & " ORDER BY " & strTablePrefix & "ANNOUNCE.A_START_DATE DESC"
strSql = strSql & ", " & strTablePrefix & "ANNOUNCE.A_ID DESC;"

set rsAnnounce =  my_Conn.Execute (strSql)

if rsAnnounce.eof or rsAnnounce.bof then
	Response.Write ""
else
	Response.Write "<table border=0 align=""center"" cellpadding=""0"" cellspacing=""0"">" & vbCrLf & _
       		       "  <tr valign=""center"">" & vbCrLf & _
		       "    <td bgcolor=""" & strTableBorderColor & """>" & vbCrLf & _
		       "      <table border=0 width=""100%"" cellspacing=""1"" cellpadding=""3"">" & vbCrLf & _
		       "        <tr>" & vbCrLf & _ 	
		       "          <td align=""center"" bgcolor=""" & strHeadCellColor & """>" & _
	               "            <font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ & color=""" & strHeadFontColor & """>��̳������</font></td>" & vbCrLf & _
		       "        </tr>" & vbCrLf & _
       		       "        <tr valign=""center"">" & vbCrLf & _
		       "          <td align=""center"" bgcolor=""" & strForumCellColor & """>" 
	while not rsAnnounce.eof 		   
	Response.Write "<font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """><a href=""javascript:openAnnounceWindow('view_announcements.asp')"">" & FormatStr(rsAnnounce("A_SUBJECT")) & "</a></font><br>"
	rsAnnounce.movenext 
	Wend
	
	Response.Write "        </td>" & vbCrLf & "</tr>" & vbCrLf & _
	               "      </table>" & vbCrLf & _
		       "    </td>" & vbCrLf & _
		       "  </tr>" & vbCrLf & _
	               "</table>" & vbCrLf
end if
%>