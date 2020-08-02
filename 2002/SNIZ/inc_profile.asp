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

 on error resume next 
 %> 

<table border="0" width="100%" cellspacing="0" cellpadding="0" valign=top>
  <tr>
    <td bgColor="<% =strPageBGColor %>" align=center  <%= strColspan %>>
    <p><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><b>前面有 <FONT color=#ff0000 size="<% =strHeaderFontSize %>"><b>*</b></FONT> 的栏目都必须填写</b></font></p>
    </td>
  </tr>
  <tr>
    <td bgcolor="<% =strPageBGColor %>" align=left valign=top>
<table border="0" width="100%" cellspacing="1" cellpadding="0" align="center">
<tr>
<%
if strUseExtendedProfile then
%>
  <td width="50%"" bgColor=<% =strPopUpTableColor %> valign=top>
    <table border="0" width="100%" cellspacing="1" cellpadding="3" bgcolor="<% =strPopUpBorderColor %>">
<tr>
<td align=center bgcolor="<% =strCategoryCellColor %>" colspan="2"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">&nbsp;联系方法&nbsp;</font></b></td>
</tr>
<tr>
<td bgColor=<% =strPopUpTableColor %> align=right width="10%"" nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><FONT color=#ff0000>*</FONT> 电子邮件：&nbsp;</font></b></td>
<td bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><INPUT  name="Email" size="25" value="<% =RS("M_EMAIL") %>">
<INPUT type="hidden" name="Email2" value="<% =RS("M_EMAIL") %>">
</tr>
<!################### - kerrycode for Email List------>
<%if Request.QueryString("mode") = "goEdit" or Request.Querystring("mode") = "goModify" then 
RECMAIL = RS("M_RECMAIL")
if RECMAIL = "1" then
YesNo = "No"
else
RECMAIL = "0"
YesNo = "Yes"
end if
%>
<tr>
<td bgColor=<% =strPopUpTableColor %> align=right width="10%"" nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">Receive Updates:&nbsp;</font></b></td>
<td bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
  <select name="RECMAIL" size="1">
  <option value="<%= RECMAIL %>" SELECTED>&nbsp;<% =YesNo%></option>
  <option value="0">&nbsp;是</option>
  <option value="1">&nbsp;否</option>
  </select>
 </font></td>
</tr>
 <% end if %>
<!################### - /kerrycode------>
<tr>
<td bgColor=<% =strPopUpTableColor %> align=right width="10%"" nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">不显示信箱地址：&nbsp;</font></td>
<td bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
<% if (RS("M_HIDE_EMAIL") = 1) then %>
是<input type="radio" value="1" checked name="HideMail">&nbsp;&nbsp;否<input type="radio" value="0" name="HideMail"
<% else %>
是<input type="radio" value="1" name="HideMail">&nbsp;&nbsp;否<input type="radio" value="0" checked name="HideMail"
<% end if %>
</font></td>


</tr>
<%if strICQ = "1" then%>
<tr>
<td bgColor=<% =strPopUpTableColor %> align=right nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">ICQ号码:&nbsp;</font></td>
<td bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><INPUT name="ICQ" size="25" value="<% =ChkString(RS("M_ICQ"), "display") %>"></font></td>
</tr>
<%end if
if strYAHOO = "1" then
%><tr>
<td bgColor=<% =strPopUpTableColor %> align=right nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">OICQ号码:&nbsp;</font></td>
<td bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><INPUT name="YAHOO" size="25" value="<% =ChkString(RS("M_YAHOO"), "display") %>"></font></td>
</tr>
<%end if
if strAIM = "1" then
%><tr>
<td bgColor=<% =strPopUpTableColor %> align=right nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">AIM号码:&nbsp;</font></td>
<td bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><INPUT name="AIM" size="25" value="<% =ChkString(RS("M_AIM"), "display") %>"></font></td>
</tr>
<%end if 
if (strHomepage + strFavLinks) > 0 and (strUseExtendedProfile) then  %>
<tr>
<td align=center bgcolor="<% =strCategoryCellColor %>" colspan="2">
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>"><b>相关连结</b>&nbsp;</font></td>
</tr>
<%if strHomepage = "1" then %>
<tr>
<td bgColor=<% =strPopUpTableColor %> align=right nowrap width="10%""><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">个人主页：&nbsp;</font></td>
<td bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><INPUT name="Homepage" size="25" value="<% if ChkString(RS("M_Homepage"), "display") <> " " and lcase(RS("M_Homepage")) <> "http://" then Response.Write(RS("M_Homepage")) else Response.Write("http://") %>"></font></td>
</tr>
<%end if
if strFavLinks = "1" then%>
<tr>
<td bgColor=<% =strPopUpTableColor %> align=right nowrap width="10%""><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">推荐连结：&nbsp;</font></td>
<td bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><INPUT name="Link1" size="25" value="<% if RS("M_LINK1") <> " " and lcase(RS("M_LINK1")) <> "http://" then Response.Write(ChkString(rs("M_LINK1"), "display")) else Response.Write("http://") %>"></font></td>
</tr>
<tr>
<td bgColor=<% =strPopUpTableColor %> align=right nowrap width="10%""><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">&nbsp;</font></td>
<td bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><INPUT name="Link2" size="25" value="<% if RS("M_LINK2") <> " " and lcase(RS("M_LINK2")) <> "http://" then Response.Write(ChkString(rs("M_LINK2"), "display")) else Response.Write("http://") %>"></font></td>
</tr>
<%end if
end if 
%>
<%if strPicture = "1" then %>
<tr>
<td align=center bgcolor="<% =strCategoryCellColor %>" colspan="2">
<b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">个性头像地址（请不要大於 70 × 70 像素）&nbsp;</font></b></td>
</tr>
<tr>
<td bgColor="<% =strPopUpTableColor %>" align=right nowrap valign=top><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">图片连接地址：&nbsp;</font></td>
<td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><INPUT  name="Photo_URL" size="25" value="<%  if Trim(RS("M_PHOTO_URL") <> "") and lcase(RS("M_PHOTO_URL")) <> "http://" and not(IsNull(RS("M_PHOTO_URL"))) then Response.Write(ChkString(rs("M_PHOTO_URL"), "display")) else Response.Write("http://") %>"></font></td>
</tr>
<%end if ' strPicture
if (strBio + strHobbies + strLNews + strQuote)> 0 then 
strMyBio = rs("M_BIO")
strMyHobbies = rs("M_HOBBIES")
strMyLNews = rs("M_LNEWS")
strMyQuote = rs("M_QUOTE")
%><tr>
<td align=center bgcolor="<% =strCategoryCellColor %>" colspan="2"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">更多详细资料</font></b></td>
</tr>
<%if strBio = "1" then  %>
<tr>
<td bgColor=<% =strPopUpTableColor %> valign=top align=right nowrap width="10%">
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">个人简介：&nbsp;</font>
</td>
<td bgColor=<% =strPopUpTableColor %> valign=top>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><textarea name="Bio" cols="30" rows=4><% =Trim(cleancode(strMyBio)) %></textarea></font>
</td>
</tr>
<%end if
if strHobbies = "1" then  %>
<tr>
<td bgColor=<% =strPopUpTableColor %> valign=top align=right nowrap width="10%">
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">个人爱好：&nbsp;</font>
</td>
<td bgColor=<% =strPopUpTableColor %> valign=top>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><textarea name="Hobbies" cols="30" rows=4><% =Trim(cleancode(strMyHobbies)) %></textarea></font>
</td>
</tr>
<%end if
if strLNews = "1" then  %>
<tr>
<td bgColor=<% =strPopUpTableColor %> valign=top align=right nowrap width="10%">
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">最近状况：&nbsp;</font>
</td>
<td bgColor=<% =strPopUpTableColor %> valign=top>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><textarea name="LNews" cols="30" rows=4><% =Trim(cleancode(strMyLNews)) %></textarea></font>
</td>
</tr>
<%end if
if strQuote = "1" then%>
<tr>
<td bgcolor="<% =strPopUpTableColor %>" valign=top align=right nowrap width="10%">
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">座 右 铭：&nbsp;</font></td>
<td bgColor=<% =strPopUpTableColor %> valign=top>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><textarea name="Quote" cols="30" rows=4><% =Trim(cleancode(strMyQuote)) %></textarea></font>
</td>
</tr>
<%end if
end if%>

<%'Rem User Field Code #######################################
if (intUserFields> 0) then 
%><tr>
<td align=center bgcolor="<% =strCategoryCellColor %>" colspan="2"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">额外资料</font></b></td>
</tr>
<tr><td></td></tr>
<%
set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open strConnString
set rs2 = Conn.Execute("SELECT * FROM " & strMemberTablePrefix & "USERFIELDS ORDER BY USR_FIELD_ID")
do until rs2.EOF
fieldValue = getUserFieldValue(rs2("USR_FIELD_ID"),RS("MEMBER_ID"))
Response.Write ("<tr><td valign=""top""  align=""right"" bgColor=" &strPopUpTableColor & " > <b><font face=""" & strDefaultFontFace & """ size=" & strDefaultFontSize & ">" & rs2("USR_LABEL") & ":</font></b></td>")
if rs2("USR_FIELDTYPE") = "T" then %>
<td valign="top" bgColor="<% =strPopUpTableColor %>"><textarea maxLength="255" name=<%=rs2("USR_SHORTNAME")%> cols="20" rows=3><%= getUserFieldValue(rs2("USR_FIELD_ID"),RS("MEMBER_ID")) %></textarea></td>
<% Else 
if rs2("USR_FIELDTYPE") = "S" then %>
<td valign="top" bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><INPUT name=<%=rs2("USR_SHORTNAME")%> size="20" value="<%= getUserFieldValue(rs2("USR_FIELD_ID"),RS("MEMBER_ID")) %>"></font></td>
<% Else  %>
<td valign="top" bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><INPUT name=<%=rs2("USR_SHORTNAME")%> type="checkbox" value="1" <% If fieldValue = "1" then %>checked<% End If %>></font></td>
<% End If %>
<% End If %>
</tr>
<%rs2.MoveNext
loop
rs2.Close
Conn.close
set Conn = Nothing
end if
'Rem User Field Code #######################################
%>
</table>
</td>
<%
end if 'extended profile
%>
<td bgColor=<% =strPopUpTableColor %> valign=top>
  <table border="0" width="100%" cellspacing="1" cellpadding="3"  bgcolor="<% =strPopUpBorderColor %>">
    <tr>
<td valign="top" align=center colspan="2" bgcolor="<% =strCategoryCellColor %>"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">基本资料</font></b></td>
</tr>
<tr>
<td bgColor=<% =strPopUpTableColor %> align=right nowrap width="10%""><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><FONT color=#ff0000>*</FONT> 用户名称：&nbsp;</font></b></td>
<td bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">
<%if (Request.QueryString("mode") = "goEdit") or (Request.Querystring("mode") = "goModify" and (Request.QueryString("Name") = "Admin") or Request.Form("MEMBER_ID") = 1 or Request.Form("Name") = "Admin") then %>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><% =ChkString(rs("M_NAME"), "display") %></font>
<INPUT type="hidden" name="Name"  value="<% =rs("M_NAME") %>">
<%else %>
<INPUT name="Name" size="25"  value="<% =ChkString(rs("M_NAME"), "display") %>">
<%end if %>
</font></td>
</tr>
<%
if Request.Querystring("mode") = "goModify" then %>
<tr>
<td bgColor="<% =strPopUpTableColor %>"align=right nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">用户名称：&nbsp;</font></b></td>
<td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><INPUT  name="Title" size="25" value="<% =CleanCode(RS("M_TITLE")) %>"></font></td>
</tr>
<%end if
if strAuthType = "nt" then %>
<tr>
<td bgColor="<% =strPopUpTableColor %>" align=right nowrap valign=top><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><FONT color=#ff0000>*</FONT>你的帐号：&nbsp;</font></b></td>
<td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSizederFontSize %>">
<%if Request.Form("Method_Type") = "Modify" then %>
<input name="Account" value="<% =ChkString(rs("M_USERNAME"), "display") %>">
<%else %>
<%=Session(strCookieURL & "userid")%><input type="hidden" name="Account" value="<% =Session(strCookieURL & "userid") %>">
<%end if %>
</font></td>
</tr>
<%else %>
<!---------------kerrycode------>
<tr>
<td bgColor="<% =strPopUpTableColor %>"align=right nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><FONT color=#ff0000>*</FONT> 用户密码：&nbsp;</font></b></td>
<td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><INPUT name="Password" type="Password" size="25" value="<% =ChkString(rs("M_PASSWORD"), "display") %>">
<INPUT name="Password-d" type=hidden value="<% =rs("M_PASSWORD") %>"></font></td>
</tr>
<%if Request.QueryString("mode") = "Register" or Request.QueryString("mode") = "goEdit" then %>
<tr>
<td bgColor="<% =strPopUpTableColor %>"align=right nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><FONT color=#ff0000>*</FONT> 确认密码：&nbsp;</font></b></td>
<td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><INPUT name="Password2" type="Password" value="<% = ChkString(rs("M_PASSWORD"), "display") %>" size="25"></font></td>
</tr>
 <!---------------/kerrycode-------->
<%end if
end if 
if strFullName = "1" then
%>
<tr>
<td bgColor="<% =strPopUpTableColor %>" align=right nowrap valign=top><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">真实名字：&nbsp;</font></td>
<td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><input name="FirstName" value="<% =rs("M_FIRSTNAME") %>"></font></td>
</tr>
<tr>
<td bgColor="<% =strPopUpTableColor %>" align=right nowrap valign=top><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">你的姓氏：&nbsp;</font></td>
<td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><input name="LastName" value="<% =rs("M_LASTNAME") %>"></font></td>
</tr>
<%
end if
if strCity = "1" then
%>
<tr>
<td bgColor="<% =strPopUpTableColor %>" align=right nowrap valign=top><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">居住城市：&nbsp;</font></td>
<td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><input name="City" value="<% =rs("M_CITY") %>"></font></td>
</tr>
<%
end if
if strState = "1" then
%>
<tr>
<td bgColor="<% =strPopUpTableColor %>" align=right nowrap valign=top><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">所在省分：&nbsp;</font></td>
<td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><input name="State" value="<% =rs("M_STATE") %>"></font></td>
</tr>
<%
end if
if strCountry = "1" then
%>
<tr>
<td bgColor="<% =strPopUpTableColor %>" align=right nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">国家地区：&nbsp;</font></td>
<td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><select name="Country" size="1">
            <OPTION selected VALUE="<% =RS("M_COUNTRY") %>"><% = ChkString(rs("M_COUNTRY"), "display") %>
                        <OPTION VALUE="Albania">Albania
                        <OPTION VALUE="Algeria">Algeria
                        <OPTION VALUE="Andorra">Andorra
                        <OPTION VALUE="Angola">Angola
                        <OPTION VALUE="Anguilla">Anguilla
                        <OPTION VALUE="Antigua and Barbuda">Antigua and Barbuda
                        <OPTION VALUE="Argentina">Argentina
                        <OPTION VALUE="Armenia">Armenia
                        <OPTION VALUE="Aruba">Aruba
                        <OPTION VALUE="Australia">Australia
                        <OPTION VALUE="Austria">Austria
                        <OPTION VALUE="Azerbaijan">Azerbaijan
                        <OPTION VALUE="Azores">Azores
                        <OPTION VALUE="Bahamas">Bahamas
                        <OPTION VALUE="Bahrain">Bahrain
                        <OPTION VALUE="Bangladesh">Bangladesh
                        <OPTION VALUE="Barbados">Barbados
<OPTION VALUE="Belarus">Belarus
                        <OPTION VALUE="Belgium">Belgium
                        <OPTION VALUE="Belize">Belize
                        <OPTION VALUE="Benin">Benin
                        <OPTION VALUE="Bermuda">Bermuda
                        <OPTION VALUE="Bhutan">Bhutan
                        <OPTION VALUE="Bolivia">Bolivia
                        <OPTION VALUE="Borneo">Borneo
                        <OPTION VALUE="Bosnia and Herzegovina">Bosnia and Herzegovina
                        <OPTION VALUE="Botswana">Botswana
                        <OPTION VALUE="Brazil">Brazil
                        <OPTION VALUE="British Indian Ocean Territories">British Indian Ocean Territories
                        <OPTION VALUE="Brunei">Brunei
                        <OPTION VALUE="Bulgaria">Bulgaria
                        <OPTION VALUE="Burkina Faso (Upper Volta)">Burkina Faso (Upper Volta)
                        <OPTION VALUE="Burundi">Burundi
                        <OPTION VALUE="Camaroon">Camaroon
                        <OPTION VALUE="Cambodia">Cambodia
                        <OPTION VALUE="加拿大">加拿大
                        <OPTION VALUE="Canary Islands">Canary Islands
                        <OPTION VALUE="Cape Vere Islands">Cape Vere Islands
                        <OPTION VALUE="Cayman Island">Cayman Island
                        <OPTION VALUE="Central African Rep">Central African Rep
                        <OPTION VALUE="Chad">Chad
                        <OPTION VALUE="Chile">Chile
                        <OPTION VALUE="中国">中国
                        <OPTION VALUE="Christmas Island">Christmas Island
                        <OPTION VALUE="Colombia">Colombia
                        <OPTION VALUE="Comoros Islands">Comoros Islands
                        <OPTION VALUE="Congo, Democratic Republic of">Congo, Democratic Republic of
                        <OPTION VALUE="Costa Rica">Costa Rica
                        <OPTION VALUE="Croatia">Croatia
                        <OPTION VALUE="Cuba">Cuba
                        <OPTION VALUE="Cyprus">Cyprus
                        <OPTION VALUE="Czech Republic">Czech Republic
                        <OPTION VALUE="Denmark">Denmark
                        <OPTION VALUE="Djibouti">Djibouti
                        <OPTION VALUE="Dominica">Dominica
                        <OPTION VALUE="Dominican Republic">Dominican Republic
                        <OPTION VALUE="East Timor">East Timor
                        <OPTION VALUE="Ecuador">Ecuador
                        <OPTION VALUE="Egypt">Egypt
                        <OPTION VALUE="El Salvador">El Salvador
                        <OPTION VALUE="Equatorial Guinea">Equatorial Guinea
                        <OPTION VALUE="Eritria">Eritria
                        <OPTION VALUE="Estonia">Estonia
                        <OPTION VALUE="Ethiopia">Ethiopia
                        <OPTION VALUE="Falkland Islands">Falkland Islands
                        <OPTION VALUE="Faroe Islands">Faroe Islands
                        <OPTION VALUE="Fed Rep Yugoslavia">Fed Rep Yugoslavia
                        <OPTION VALUE="Fiji">Fiji
                        <OPTION VALUE="Finland">Finland
                        <OPTION VALUE="France">France
                        <OPTION VALUE="French Guiana">French Guiana
                        <OPTION VALUE="French Polynesia">French Polynesia
                        <OPTION VALUE="Fyro Macedonia">Fyro Macedonia
                        <OPTION VALUE="Gabon">Gabon
                        <OPTION VALUE="Gambia">Gambia
                        <OPTION VALUE="Georgia">Georgia
                        <OPTION VALUE="Germany">Germany
                        <OPTION VALUE="Ghana">Ghana
                        <OPTION VALUE="Gibraltar">Gibraltar
                        <OPTION VALUE="Greece">Greece
                        <OPTION VALUE="Greenland">Greenland
                        <OPTION VALUE="Grenada">Grenada
                        <OPTION VALUE="Guadeloupe">Guadeloupe
                        <OPTION VALUE="Guatemala">Guatemala
                        <OPTION VALUE="Guinea">Guinea
                        <OPTION VALUE="Guinea-Bissau">Guinea-Bissau
                        <OPTION VALUE="Guyana">Guyana
                        <OPTION VALUE="Haiti">Haiti
                        <OPTION VALUE="Honduras">Honduras
                        <OPTION VALUE="香港">香港
                        <OPTION VALUE="Hungary">Hungary
                        <OPTION VALUE="Iceland">Iceland
                        <OPTION VALUE="India">India
                        <OPTION VALUE="Indonesia">Indonesia
                        <OPTION VALUE="Iran">Iran
                        <OPTION VALUE="Iraq">Iraq
                        <OPTION VALUE="Ireland">Ireland
                        <OPTION VALUE="Israel">Israel
                        <OPTION VALUE="Italy">Italy
                        <OPTION VALUE="Ivory Coast">Ivory Coast
                        <OPTION VALUE="Jamaica">Jamaica
                        <OPTION VALUE="日本">日本
                        <OPTION VALUE="Jordan">Jordan
                        <OPTION VALUE="Kazakhstan">Kazakhstan
                        <OPTION VALUE="Kenya">Kenya
                        <OPTION VALUE="Kiribati">Kiribati
                        <OPTION VALUE="Korea">Korea
                        <OPTION VALUE="Kuwait">Kuwait
                        <OPTION VALUE="Kyrgyzstan">Kyrgyzstan
                        <OPTION VALUE="Laos">Laos
                        <OPTION VALUE="Latvia">Latvia
                        <OPTION VALUE="Lebanon">Lebanon
                        <OPTION VALUE="Lesotho">Lesotho
                        <OPTION VALUE="Liberia">Liberia
                        <OPTION VALUE="Libya">Libya
                        <OPTION VALUE="Liechtenstein">Liechtenstein
                        <OPTION VALUE="Lithuania">Lithuania
                        <OPTION VALUE="Luxembourg">Luxembourg
                        <OPTION VALUE="Macao">Macao
                        <OPTION VALUE="Madagascar">Madagascar
                        <OPTION VALUE="Malawi">Malawi
                        <OPTION VALUE="Malaysia">Malaysia
                        <OPTION VALUE="Maldives">Maldives
                        <OPTION VALUE="Mali">Mali
                        <OPTION VALUE="Malta">Malta
                        <OPTION VALUE="Martinique">Martinique
                        <OPTION VALUE="Mauritania">Mauritania
                        <OPTION VALUE="Mauritius">Mauritius
                        <OPTION VALUE="Mexico">Mexico
                        <OPTION VALUE="Moldova">Moldova
                        <OPTION VALUE="Monaco">Monaco
                        <OPTION VALUE="Mongolia">Mongolia
                        <OPTION VALUE="Montserrat">Montserrat
                        <OPTION VALUE="Morocco">Morocco
                        <OPTION VALUE="Mozambique">Mozambique
                        <OPTION VALUE="Myanmar (Burma)">Myanmar (Burma)
                        <OPTION VALUE="Namibia">Namibia
                        <OPTION VALUE="Naura">Naura
                        <OPTION VALUE="Nepal">Nepal
                        <OPTION VALUE="Netherlands">Netherlands
                        <OPTION VALUE="Netherlands Antilles">Netherlands Antilles
                        <OPTION VALUE="New Caledonia">New Caledonia
                        <OPTION VALUE="New Zealand">New Zealand
                        <OPTION VALUE="Nicaragua">Nicaragua
                        <OPTION VALUE="Niger">Niger
                        <OPTION VALUE="Nigeria">Nigeria
                        <OPTION VALUE="Niue">Niue
                        <OPTION VALUE="Norway">Norway
                        <OPTION VALUE="Oman">Oman
                        <OPTION VALUE="Pakistan">Pakistan
            <OPTION VALUE="Palestine">Palestine
                        <OPTION VALUE="Panama">Panama
                        <OPTION VALUE="Papua New Guinea">Papua New Guinea
                        <OPTION VALUE="Paraguay">Paraguay
                        <OPTION VALUE="Peru">Peru
                        <OPTION VALUE="Philippines">Philippines
                        <OPTION VALUE="Pitcairn Island">Pitcairn Island
                        <OPTION VALUE="Poland">Poland
                        <OPTION VALUE="Portugal">Portugal
                        <OPTION VALUE="Qatar">Qatar
                        <OPTION VALUE="Republic of Korea">Republic of Korea
                        <OPTION VALUE="Reunion Island">Reunion Island
                        <OPTION VALUE="Romania">Romania
                        <OPTION VALUE="Russia">Russia
                        <OPTION VALUE="Rwanda">Rwanda
                        <OPTION VALUE="Saint Barthelemy">Saint Barthelemy
                        <OPTION VALUE="Saint Croix">Saint Croix
                        <OPTION VALUE="Saint Helena">Saint Helena
                        <OPTION VALUE="Saint Kitts and Nevis">Saint Kitts and Nevis
                        <OPTION VALUE="Saint Lucia">Saint Lucia
                        <OPTION VALUE="Saint Pierre and Miquelon">Saint Pierre and Miquelon
                        <OPTION VALUE="Saint Vincent and Grenadi">Saint Vincent and Grenadi
                        <OPTION VALUE="San Marino">San Marino
                        <OPTION VALUE="Sao Tome and Principe">Sao Tome and Principe
                        <OPTION VALUE="Saudi Arabia">Saudi Arabia
                        <OPTION VALUE="Senegal">Senegal
                        <OPTION VALUE="Seychelles">Seychelles
                        <OPTION VALUE="Sierra Leone">Sierra Leone
                        <OPTION VALUE="Singapore">Singapore
                        <OPTION VALUE="Slovakia">Slovakia
                        <OPTION VALUE="Slovenia">Slovenia
                        <OPTION VALUE="Solomon Islands">Solomon Islands
                        <OPTION VALUE="Somalia Northern Region">Somalia Northern Region
                        <OPTION VALUE="Somalia Southern Region">Somalia Southern Region
                        <OPTION VALUE="South Africa">South Africa
                        <OPTION VALUE="South Sandwich Islands">South Sandwich Islands
                        <OPTION VALUE="Spain">Spain
                        <OPTION VALUE="Sri Lanka">Sri Lanka
                        <OPTION VALUE="Sudan">Sudan
                        <OPTION VALUE="Suriname">Suriname
                        <OPTION VALUE="Swaziland">Swaziland
                        <OPTION VALUE="Sweden">Sweden
                        <OPTION VALUE="Switzerland">Switzerland
                        <OPTION VALUE="Syria">Syria
                        <OPTION VALUE="中国" selected>中国
                        <OPTION VALUE="Tajikistan">Tajikistan
                        <OPTION VALUE="Tanzania">Tanzania
                        <OPTION VALUE="Thailand">Thailand
                        <OPTION VALUE="Togo">Togo
                        <OPTION VALUE="Tonga">Tonga
                        <OPTION VALUE="Trinidad and Tobago">Trinidad and Tobago
                        <OPTION VALUE="Tunisia">Tunisia
                        <OPTION VALUE="Turkey">Turkey
                        <OPTION VALUE="Turkmenistan">Turkmenistan
                        <OPTION VALUE="Turks and Caicos Islnd">Turks and Caicos Islnd
                        <OPTION VALUE="Tuvalu">Tuvalu
                        <OPTION VALUE="美国">美国
                        <OPTION VALUE="Uganda">Uganda
                        <OPTION VALUE="Ukraine">Ukraine
                        <OPTION VALUE="United Arab Emirates">United Arab Emirates
                        <OPTION VALUE="United Kingdom">United Kingdom
                        <OPTION VALUE="Uruguay">Uruguay
                        <OPTION VALUE="Uzbekistan">Uzbekistan
                        <OPTION VALUE="Vanuatu">Vanuatu
                        <OPTION VALUE="Vatican City">Vatican City
                        <OPTION VALUE="Venezuela">Venezuela
                        <OPTION VALUE="Vietnam">Vietnam
                        <OPTION VALUE="Virgin Islands (United Kingdom)">Virgin Islands (United Kingdom)
                        <OPTION VALUE="Wallis and Futuna Islands">Wallis and Futuna Islands
                        <OPTION VALUE="Western Sahara">Western Sahara
                        <OPTION VALUE="Western Samoa">Western Samoa
                        <OPTION VALUE="Yemen">Yemen
<OPTION VALUE="Zambia">Zambia
<OPTION VALUE="Zimbabwe (Rhodesia)">Zimbabwe (Rhodesia)         
</select></font></td>
</tr>
<%end if
if strAge = "1" then
%>
<tr>
<td bgColor="<% =strPopUpTableColor %>" align=right nowrap valign=top><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">年&nbsp;&nbsp;&nbsp;&nbsp;龄：&nbsp;</font></td>
<td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><INPUT  name="Age" size="7" value="<% =ChkString(RS("M_AGE"), "display") %>"></font></td>
</tr>
<%
end if
if strSex = "1" then
%>
<tr>
<td bgColor="<% =strPopUpTableColor %>" align=right nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">性&nbsp;&nbsp;&nbsp;&nbsp;别：&nbsp;</font>></td>
<td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><select name="Sex" size="1">
            <OPTION selected VALUE="<% =RS("M_SEX") %>"><% =RS("M_SEX") %>&nbsp;
            <OPTION VALUE="男性">男性&nbsp;
            <OPTION VALUE="女性">女性&nbsp;
            </select></font></td>
</tr>
<%
end if
if strMarStatus = "1" then
%>
<tr>
<td bgColor="<% =strPopUpTableColor %>" align=right nowrap valign=top><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">婚姻状况：&nbsp;</font></td>
<td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><INPUT  name="MarStatus" size="25" value="<% =ChkString(RS("M_MARSTATUS"), "display") %>"></font></td>
</tr>
<%
end if
if strOccupation = "1" then
%><tr>
<td bgColor="<% =strPopUpTableColor %>" align=right nowrap valign=top><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">现任职业：&nbsp;</font></td>
<td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><INPUT  name="Occupation" size="25" value="<% = ChkString(RS("M_OCCUPATION"), "display") %>"></font></td>
</tr>
<%
end if
if Request.Querystring("mode") = "goModify" then %>
<tr>
<td bgColor="<% =strPopUpTableColor %>"align=right nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">发表总数：&nbsp;</font></td>
<td bgColor="<% =strPopUpTableColor %>"><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><INPUT  name="Posts" size="25" value="<% = ChkString(RS("M_POSTS"), "display") %>"></font></td>
</tr>
<%end if
strTxtSig = RS("M_SIG") %>
<tr>
<td bgColor="<% =strPopUpTableColor %>" align=right valign=top nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">个人签名：&nbsp;</font></td>
<td bgColor="<% =strPopUpTableColor %>"><textarea maxLength="255" name="Sig" cols="25" rows=4><% =Trim(cleancode(strTxtSig)) %></textarea></td>
</tr>
<%if Request.Form("Method_Type") = "Modify" then %>
<tr>
<td bgColor="<% =strPopUpTableColor %>" align=right valign=top nowrap><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">会员等级：&nbsp;</font></td>
<td bgColor="<% =strPopUpTableColor %>" valign=top>
<%if rs("M_NAME") = "Admin" then %>
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">管理者</font>
<input type="hidden" value="3" name="Level">
<%else %>
        <SELECT value="1" name="Level">
<OPTION VALUE="1"<% if rs("M_LEVEL") = 1 then Response.Write(" selected") %>>一般使用者
<OPTION VALUE="2"<% if rs("M_LEVEL") = 2 then Response.Write(" selected") %>>版主
<OPTION VALUE="3"<% if rs("M_LEVEL") = 3 then Response.Write(" selected") %>>管理者
</SELECT>
<%end if %>
</td>
</tr>
<% '####################################File Attachment ######################### %>
<tr>
<td bgColor="<% =strPopUpTableColor %>" align=right valign=top nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">附加文件：&nbsp;</font></b></td>
<td bgColor="<% =strPopUpTableColor %>" valign=top>
<input type="checkbox" name="allowDownloads" value="1"<% if rs("M_ALLOWDOWNLOADS") = 1 then Response.Write(" checked") end if%> > <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">允许下载</font><br>
<input type="checkbox" name="allowUploads" value="1"<% if rs("M_ALLOWUPLOADS") = 1 then Response.Write(" checked") end if%> > <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">允许上传</font>
</td>
</tr>
<% '####################################File Attachment ######################### %>
<%end if
if not(strUseExtendedProfile)  then
%>
 <tr>
<td align=center bgcolor="<% =strCategoryCellColor %>" colspan="2"><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">&nbsp;联系方法&nbsp;</font></b></td>
 </tr>
 <tr>
<td bgColor=<% =strPopUpTableColor %> align=right width="10%"" nowrap><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><FONT color=#ff0000>*</FONT> 电子邮件：&nbsp;</font></b></td>
<td bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><INPUT  name="Email" size="25" value="<% = ChkString(RS("M_EMAIL"), "display") %>">
<INPUT type="hidden" name="Email2" value="<% =RS("M_EMAIL") %>"></font></td>
</tr>
<%if strICQ = "1" then%>
<tr>
<td bgColor=<% =strPopUpTableColor %> align=right nowrap width="10%""><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">ICQ号码:&nbsp;</font></b></td>
<td bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><INPUT name="ICQ" size="25" value="<% = ChkString(RS("M_ICQ"), "display") %>"></font></td>
</tr>
<%end if
if strYAHOO = "1" then
%><tr>
<td bgColor=<% =strPopUpTableColor %> align=right nowrap width="10%""><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">YAHOO号码:&nbsp;</font></b></td>
<td bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><INPUT name="YAHOO" size="25" value="<% = ChkString(RS("M_YAHOO"), "display") %>"></font></td>
</tr>
<%end if
if strAIM = "1" then
%><tr>
<td bgColor=<% =strPopUpTableColor %> align=right nowrap width="10%""><b><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">OICQ号码:&nbsp;</font></b></td>
<td bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><INPUT name="AIM" size="25" value="<% = ChkString(RS("M_AIM"), "display") %>"></font></td>
</tr>
<%end if 
end if
if (strHomepage + strFavLinks) > 0 and not(strUseExtendedProfile) then  %>
<tr>
<td align=center bgcolor="<% =strCategoryCellColor %>" colspan="2">
<font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strHeadFontColor %>">相关连结&nbsp;</font></td>
</tr>
<%if strHomepage = "1" then %>
<tr>
<td bgColor=<% =strPopUpTableColor %> align=right nowrap width="10%""><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">个人主页：&nbsp;</font></td>
<td bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><INPUT name="Homepage" size="25" value="<% if RS("M_Homepage") <> " " and lcase(RS("M_Homepage")) <> "http://" then Response.Write(ChkString(RS("M_Homepage"), "display")) else Response.Write("http://") %>"></font></td>
</tr>
<%end if
if strFavLinks = "1" then
 strLINK1 = rs(M_LINK1)
 strLINK2 = rs(M_LINK2)
%>
<tr>
<td bgColor=<% =strPopUpTableColor %> align=right nowrap width="10%""><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">推荐连结：&nbsp;</font></td>
<td bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><INPUT name="Link1" size="25" value="<% if strLINK1 <> " " and strLINK1 <> "http://" then Response.Write(ChkString(strLINK1, "display")) else Response.Write("http://") %>"></font></td>
</tr>
<tr>
<td bgColor=<% =strPopUpTableColor %> align=right nowrap width="10%""><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>">&nbsp;</font></td>
<td bgColor=<% =strPopUpTableColor %>><font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><INPUT name="Link2" size="25" value="<% if strLINK2 <> " " and lcase(strLINK2) <> "http://" then Response.Write(ChkString(strLINK2, "display")) else Response.Write("http://") %>"></font></td>
</tr>
<%end if
end if
%>
</table>
</td>
</tr>
</table>
</td>
  </tr>
</table>
    </td>
  </tr>
<% if strUseExtendedProfile then %>  
</table>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td bgColor="<% =strPageBGColor %>" align=center nowrap <%= strColspan %>>
    <br>      
        <INPUT type="hidden" value="<% =Request.Form("MEMBER_ID") %>" name="MEMBER_ID">
        <INPUT type="submit" value="  提  交  " name=Submit1>
      </td>
  </tr>
<% else%>
  <tr>
    <td bgColor="<% =strPageBGColor %>" align=center nowrap <%= strColspan %>>
    <br>      
        <INPUT type="hidden" value="<% =Request.Form("MEMBER_ID") %>" name="MEMBER_ID">
        <INPUT type="submit" value="马上提交注册信息" name=Submit1>
      </td>
  </tr>
<% end if%>
</table>