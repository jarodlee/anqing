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
<!--#include file="config.asp"-->


<!--#include file="inc_pollmentor.asp"-->
<!--#INCLUDE FILE="inc_functions.asp" -->
<!--#INCLUDE FILE="inc_top.asp" -->
<%
Dim id
Dim count1, answer1
Dim count2, answer2
Dim count3, answer3
Dim count4, answer4
Dim count5, answer5
Dim count6, answer6
Dim count7, answer7
Dim count8, answer8


count1 = Request.Form("count1")
answer1 = Request.Form("answer1")
count2 = Request.Form("count2")
answer2 = Request.Form("answer2")
count3 = Request.Form("count3")
answer3 = Request.Form("answer3")
count4 = Request.Form("count4")
answer4 = Request.Form("answer4")
count5 = Request.Form("count5")
answer5 = Request.Form("answer5")
count6 = Request.Form("count6")
answer6 = Request.Form("answer6")
count7 = Request.Form("count7")
answer7 = Request.Form("answer7")
count8 = Request.Form("count8")
answer8 = Request.Form("answer8")

	Dim oConn
    Set oConn = PollMentor_GetDatabaseConn()
    Dim oRS
    
    
	If Request.QueryString("save")="yes" Then
		Set oRS = Server.CreateObject("ADODB.Recordset")
		
		Dim sID
		sID = Request.QueryString("id")
		If sID = "" Then
			sID = 0
		Else
			sID = CInt( sID)
		End If
		
		oRS.Open "select * from " & StrTablePrefix& "poll where id = " & sID, oConn, 1, 3
		
		If Request.QueryString("action") = "new" Or Request.QueryString("action") ="edit" Then
			If Request.QueryString("action") = "new"  Then
				oRS.AddNew
			End If
			If count1 = "" Then
				count1=0
			End If
			oRS("count1") = count1
			oRS("answer1") = answer1
			If count2 = "" Then
				count2=0
			End If
			oRS("count2") = count2
			oRS("answer2") = answer2
			If count3 = "" Then
				count3=0
			End If
			oRS("count3") = count3
			oRS("answer3") = answer3
			If count4 = "" Then
				count4=0
			End If
			oRS("count4") = count4
			oRS("answer4") = answer4
			If count5 = "" Then
				count5=0
			End If
			oRS("count5") = count5
			oRS("answer5") = answer5
			If count6 = "" Then
				count6=0
			End If
			oRS("count6") = count6
			oRS("answer6") = answer6
			If count7 = "" Then
				count7=0
			End If
			oRS("count7") = count7
			oRS("answer7") = answer7
			If count8 = "" Then
				count8=0
			End If
			oRS("count8") = count8
			oRS("answer8") = answer8
			oRS("question") = Request.Form("question")
			oRS("createdwhen") = Now()
		End If
		If Request.QueryString("action") = "del"  Then
				oRS.Delete
		Else
			oRS.Update
		End If
		oRS.Close
	Set oRS = Nothing
	oConn.Close
	Set oConn = Nothing
	Response.Redirect "admin_pollmentor.asp"
	End If
'''''''''''  (C) Stefan Holmberg 1999  
'''''''''''  Free to use if these sourcecode lines is not deleted 
'''''''''''  Contact me at webmaster@sqlexperts.com
'''''''''''  http://www.sqlexperts.com
'''''''''''  AdMentor homepage at http://www.create-a-webshop.com
If Request.QueryString("action") = "edit" Then
	Dim oRS2
	Set oRS2 = oConn.Execute( "select * from "  & strTablePrefix  & "poll where id = " & Request.QueryString("id") )
	question = oRS2("question")
	answer1 = oRS2("answer1")
	count1 = oRS2("count1")
	answer2 = oRS2("answer2")
	count2 = oRS2("count2")
	answer3 = oRS2("answer3")
	count3 = oRS2("count3")
	answer4 = oRS2("answer4")
	count4 = oRS2("count4")
	answer5 = oRS2("answer5")
	count5 = oRS2("count5")
	answer6 = oRS2("answer6")
	count6 = oRS2("count6")
	answer7 = oRS2("answer7")
	count7 = oRS2("count7")
	answer8 = oRS2("answer8")
	count8 = oRS2("count8")
	
	oRS2.Close
Else
End If
%>



<table align="center"  border="0" cellPadding="3" cellSpacing="0" width="80%">
  <tbody>
    <tr>
      <td vAlign="top"  height="40">
      </td>
    </tr>
    <tr>
      <td vAlign="top" colspan="2">
        <table align="center" bgcolor="<% =strTableBorderColor %>" border="0" cellPadding="0" cellSpacing="1">
          <tbody>
            <tr>
              <td height="100%" vAlign="top" width="85%">
                <table bgcolor="<% =strTableBorderColor %>" border="0" cellPadding="3" cellSpacing="1">
                  <tbody>
                    <tr bgcolor="<% =strForumCellColor %>">
                      <td align="left" height="100%" vAlign="top" width="65%">
                        <table border="0" width="100%">
                          <tr>
                            <td width="50%"><font color="#aa3333" face="verdana,arial,helvetica" size="4">ͶƱ���� <%=Session("fullname")%></font>
                             
                            </td>
                            <td width="50%">
                            <%=FAQ_GetAd(1)
                            %>
                            </td>
                          </tr>
                        </table>
                        <hr color="#000066" noShade SIZE="1">
                             
                             
                        <table border="0" width="100%">
                          <tr>
                            <td width="70%">
                             
                        <p align="left"><b>Question</b><br>
                        <%
                        Dim sURL
                        sURL = "admin_poll.asp?save=yes&action=" & Request.QueryString("action")
                       	sURL = sURL & "&id=" & Request.QueryString("id") 
                        %>
                        <form method="POST" action="<%=sURL%>">
                          <table border="0" width="100%" bgcolor="<% =strTableBorderColor %>" cellspacing="1">
                            <tr bgcolor="<% =strForumCellColor %>">
                              <td width="29%">ͶƱ����:</td>
                              <td width="71%"><input type="text" name="question" size="40" value="<%=question%>"></td>
                            </tr>
                            <tr bgcolor="<% =strForumCellColor %>">
                              <td width="29%">ѡ�� 1:<br>
 <input type="text" name="answer1" size="20" value="<%=answer1%>"></td>
                              <td width="71%">Nr of votes:<br>
                                <input type="text" name="count1" size="20" value="<%=count1%>"></td>
                            </tr>
                            <tr bgcolor="<% =strForumCellColor %>">
                              <td width="29%">ѡ�� 2:<br>
 <input type="text" name="answer2" size="20" value="<%=answer2%>"></td>
                              <td width="71%">Nr of votes:<br>
                                <input type="text" name="count2" size="20" value="<%=count2%>"></td>
                            </tr>
                            <tr bgcolor="<% =strForumCellColor %>">
                              <td width="29%">ѡ�� 3:<br>
 <input type="text" name="answer3" size="20" value="<%=answer3%>"></td>
                              <td width="71%">Nr of votes:<br>
                                <input type="text" name="count3" size="20" value="<%=count3%>"></td>
                            </tr>
                            <tr bgcolor="<% =strForumCellColor %>">
                              <td width="29%">ѡ�� 4:<br>
 <input type="text" name="answer4" size="20" value="<%=answer4%>"></td>
                              <td width="71%">Nr of votes:<br>
                                <input type="text" name="count4" size="20" value="<%=count4%>"></td>
                            </tr>
                            <tr bgcolor="<% =strForumCellColor %>">
                              <td width="29%">ѡ�� 5:<br>
 <input type="text" name="answer5" size="20" value="<%=answer5%>"></td>
                              <td width="71%">Nr of votes:<br>
                                <input type="text" name="count5" size="20" value="<%=count5%>"></td>
                            </tr>
                            <tr bgcolor="<% =strForumCellColor %>">
                              <td width="29%">ѡ�� 6:<br>
 <input type="text" name="answer6" size="20" value="<%=answer6%>"></td>
                              <td width="71%">Nr of votes:<br>
                                <input type="text" name="count6" size="20" value="<%=count6%>"></td>
                            </tr>
                            <tr bgcolor="<% =strForumCellColor %>">
                              <td width="29%">ѡ�� 7:<br>
 <input type="text" name="answer7" size="20" value="<%=answer7%>"></td>
                              <td width="71%">Nr of votes:<br>
                                <input type="text" name="count7" size="20" value="<%=count7%>"></td>
                            </tr>
                            <tr bgcolor="<% =strForumCellColor %>">
                              <td width="29%">ѡ�� 8:<br>
 <input type="text" name="answer8" size="20" value="<%=answer8%>"></td>
                              <td width="71%">Nr of votes:<br>
                                <input type="text" name="count8" size="20" value="<%=count8%>"></td>
                            </tr>
                          </table>
                          <p align="left"><input type="submit" value="�ύ" name="B1"></p>
                        </form>
                            
                        <br>
                             
                        <table border="0" width="100%">
                          <tr>
                            <td width="50%">
                             <%=FAQ_GetAd(3)
                             %>
                            </td>
                            <td width="50%">
                            <%=FAQ_GetAd(2)
                            %>
                            </td>
                          </tr>
                        </table>
                          </tr>
                          <tr>
                            <td width="70%">
                             
                            </td>
                          </tr>
                        </table>
                      </td>
                    </tr>
                  </tbody>
                </table>
              </td>
            </tr>
          </tbody>
        </table>
      </td>
    </tr>
  </tbody>
</table>
<div align="center"><a href="admin_pollmentor.asp"><font face="<% =strDefaultFontFace %>" color="<% =strForumFontColor %>" size="<% =strDefaultFontSize %>">����</font></a></div>
<!--#include file="inc_footer.asp"-->