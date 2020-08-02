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

function ChkUrls(fString, fTestTag, fType)

Dim strArray
Dim Counter
Dim strTempString

strTempString = fString
if Instr(1, fString, fTestTag) > 0 then
	strArray = Split(fString, fTestTag, -1)
	strTempString = strArray(0)

	for counter = 1 to UBound(strArray)
		if ((strArray(counter-1) = "" or len(strArray(counter-1)) < 5) and strArray(counter)<> "") then
			strTempString = strTempString & edit_hrefs("" & fTestTag & strArray(counter), fType)
		elseif ((UCase(right(strArray(counter-1),6)) <> "HREF=""") and (UCase(right(strArray(counter-1),5)) <> "[URL]") and (UCase(right(strArray(counter-1),6)) <> "[URL=""") and (UCase(right(strArray(counter-1),7)) <> "FILE:///") and (UCase(right(strArray(counter-1),7)) <> "HTTP://") and (UCase(right(strArray(counter-1),8)) <> "HTTPS://") and (UCase(right(strArray(counter-1),5)) <> "SRC=""") and (UCase(right(strArray(counter-1),5)) <> "SRC=""") and strArray(counter)<> "") then
			strTempString = strTempString & edit_hrefs("" & fTestTag & strArray(counter), fType)
		else
			strTempString = strTempString & fTestTag & strArray(counter)
		end if
	next


	

   
end if

ChkUrls = strTempString

end function


function ChkMail(fString, fTestTag, fType)

Dim strArray
Dim Counter
Dim strTempString

strTempString = fString

if Instr(1, fString, fTestTag) > 0 then
	strArray = Split(fString, fTestTag, -1)
	strTempString = ""
'	strTempString = strArray(0)
	for counter = 0 to UBound(strArray)
		if (Instr(strArray(counter), "@") > 0) and not(Instr(strArray(counter), "mailto:") > 0) and not(Instr(UCase(strArray(counter)), "[URL") > 0) then
			strTempString = strTempString & edit_hrefs("" & fTestTag & strArray(counter), fType)
		else
			strTempString = strTempString & fTestTag & strArray(counter)
		end if
	next

end if

ChkMail = strTempString

end function


function FormatStr(fString)
	on Error resume next
	fString = Replace(fString, CHR(13), "")
	fString = Replace(fString, CHR(10) & CHR(10), "</P><P>")
	fString = Replace(fString, CHR(10), "<BR>")
	if strBadWordFilter = 1 then
		fString = ChkBadWords(fString)
	end if
	fString = ChkUrls(fString,"http://", 1)
	fString = ChkUrls(fString,"https://", 2)
	fString = ChkUrls(fString,"file:///", 3)
	fString = ChkUrls(fString,"www.", 4)
	fString = ChkUrls(fString,"mailto:",5)
	fString = ChkMail(fString," ",5)

	'fString = edit_hrefs(fString, 5)
	fString = ReplaceUrls(fString)
	FormatStr = fString
end function

function doublenum(fNum)
	if fNum > 9 then 
		doublenum = fNum 
	else 
		doublenum = "0" & fNum
	end if
end function

function widenum(fNum)
	if fNum > 9 then 
		widenum = "" 
	else 
		widenum = "&nbsp;"
	end if
end function

function Chked(fYN)
   if fYN = "yes" or fYN = "1" or fYN = 1 then '**
      Chked = " Checked"
   else 
      Chked = ""
   end if    
end function

function doCode(fString, fOTag, fCTag, fROTag, fRCTag)
	fOTagPos = Instr(1, fString, fOTag, 1)
	fCTagPos = Instr(1, fString, fCTag, 1)
	while (fCTagPos > 0 and fOTagPos > 0)
		fString = replace(fString, fOTag, fROTag, 1, 1, 1)
		fString = replace(fString, fCTag, fRCTag, 1, 1, 1)
		fOTagPos = Instr(1, fString, fOTag, 1)
		fCTagPos = Instr(1, fString, fCTag, 1)
	wend
	doCode = fString
end function

function CleanCode(fString)
	if fString = "" then 
		fString = " "
	else 
		if strAllowForumCode = "1" then
			fString = replace(fString, "<b>","[b]", 1, -1, 1)
			fString = replace(fString, "</b>","[/b]", 1, -1, 1)
		    fString = replace(fString, "<s>", "[s]", 1, -1, 1)
		    fString = replace(fString, "</s>", "[/s]", 1, -1, 1)
			fString = replace(fString, "<u>","[u]", 1, -1, 1)
			fString = replace(fString, "</u>","[/u]", 1, -1, 1)
			fString = replace(fString, "<i>","[i]", 1, -1, 1)
			fString = replace(fString, "</i>","[/i]", 1, -1, 1)
			fString = replace(fString, "<font face='Andale Mono'>", "[font=Andale Mono]", 1, -1, 1)
			fString = replace(fString, "</font id='Andale Mono'>", "[/font=Andale Mono]", 1, -1, 1)
			fString = replace(fString, "<font face='Arial'>", "[font=Arial]", 1, -1, 1)
			fString = replace(fString, "</font id='Arial'>", "[/font=Arial]", 1, -1, 1)
			fString = replace(fString, "<font face='Arial Black'>", "[font=Arial Black]", 1, -1, 1)
			fString = replace(fString, "</font id='Arial Black'>", "[/font=Arial Black]", 1, -1, 1)
			fString = replace(fString, "<font face='Book Antiqua'>", "[font=Book Antiqua]", 1, -1, 1)
			fString = replace(fString, "</font id='Book Antiqua'>", "[/font=Book Antiqua]", 1, -1, 1)
			fString = replace(fString, "<font face='Century Gothic'>", "[font=Century Gothic]", 1, -1, 1)
			fString = replace(fString, "</font id='Century Gothic'>", "[/font=Century Gothic]", 1, -1, 1)
			fString = replace(fString, "<font face='Comic Sans MS'>", "[font=Comic Sans MS]", 1, -1, 1)
			fString = replace(fString, "</font id='Comic Sans MS'>", "[/font=Comic Sans MS]", 1, -1, 1)
			fString = replace(fString, "<font face='Courier New'>", "[font=Courier New]", 1, -1, 1)
			fString = replace(fString, "</font id='Courier New'>", "[/font=Courier New]", 1, -1, 1)
			fString = replace(fString, "<font face='Georgia'>", "[font=Georgia]", 1, -1, 1)
			fString = replace(fString, "</font id='Georgia'>", "[/font=Georgia]", 1, -1, 1)
			fString = replace(fString, "<font face='Impact'>", "[font=Impact]", 1, -1, 1)
			fString = replace(fString, "</font id='Impact'>", "[/font=Impact]", 1, -1, 1)
			fString = replace(fString, "<font face='Tahoma'>", "[font=Tahoma]", 1, -1, 1)
			fString = replace(fString, "</font id='Tahoma'>", "[/font=Tahoma]", 1, -1, 1)
			fString = replace(fString, "<font face='Times New Roman'>", "[font=Times New Roman]", 1, -1, 1)
			fString = replace(fString, "</font id='Times New Roman'>", "[/font=Times New Roman]", 1, -1, 1)
			fString = replace(fString, "<font face='Trebuchet MS'>", "[font=Trebuchet MS]", 1, -1, 1)
			fString = replace(fString, "</font id='Trebuchet MS'>", "[/font=Trebuchet MS]", 1, -1, 1)
			fString = replace(fString, "<font face='Script MT Bold'>", "[font=Script MT Bold]", 1, -1, 1)
			fString = replace(fString, "</font id='Script MT Bold'>", "[/font=Script MT Bold]", 1, -1, 1)
			fString = replace(fString, "<font face='Stencil'>", "[font=Stencil]", 1, -1, 1)
			fString = replace(fString, "</font id='Stencil'>", "[/font=Stencil]", 1, -1, 1)
			fString = replace(fString, "<font face='宋体'>", "[font=宋体]", 1, -1, 1)
			fString = replace(fString, "</font id='宋体'>", "[/font=宋体]", 1, -1, 1)
			fString = replace(fString, "<font face='Lucida Console'>", "[font=Lucida Console]", 1, -1, 1)
			fString = replace(fString, "</font id='Lucida Console'>", "[/font=Lucida Console]", 1, -1, 1)
		    
		      fString = replace(fString, "<font color=red>", "[red]", 1, -1, 1)
		      fString = replace(fString, "</font id=red>", "[/red]", 1, -1, 1)
		      fString = replace(fString, "<font color=green>", "[green]", 1, -1, 1)
		      fString = replace(fString, "</font id=green>", "[/green]", 1, -1, 1)
		      fString = replace(fString, "<font color=blue>", "[blue]", 1, -1, 1)
		      fString = replace(fString, "</font id=blue>", "[/blue]", 1, -1, 1)
		      fString = replace(fString, "<font color=white>", "[white]", 1, -1, 1)
		      fString = replace(fString, "</font id=white>", "[/white]", 1, -1, 1)
		      fString = replace(fString, "<font color=purple>", "[purple]", 1, -1, 1)
		      fString = replace(fString, "</font id=purple>", "[/purple]", 1, -1, 1)
	  	      fString = replace(fString, "<font color=yellow>", "[yellow]", 1, -1, 1)
	  	      fString = replace(fString, "</font id=yellow>", "[/yellow]", 1, -1, 1)
		      fString = replace(fString, "<font color=violet>", "[violet]", 1, -1, 1)
		      fString = replace(fString, "</font id=violet>", "[/violet]", 1, -1, 1)
		      fString = replace(fString, "<font color=brown>", "[brown]", 1, -1, 1)
		      fString = replace(fString, "</font id=brown>", "[/brown]", 1, -1, 1)
		      fString = replace(fString, "<font color=black>", "[black]", 1, -1, 1)
		      fString = replace(fString, "</font id=black>", "[/black]", 1, -1, 1)
		      fString = replace(fString, "<font color=pink>", "[pink]", 1, -1, 1)
		      fString = replace(fString, "</font id=pink>", "[/pink]", 1, -1, 1)
		      fString = replace(fString, "<font color=orange>", "[orange]", 1, -1, 1)
		      fString = replace(fString, "</font id=orange>", "[/orange]", 1, -1, 1)
		      fString = replace(fString, "<font color=gold>", "[gold]", 1, -1, 1)
		      fString = replace(fString, "</font id=gold>", "[/gold]", 1, -1, 1)

		      fString = replace(fString, "<font color=beige>", "[beige]", 1, -1, 1)
		      fString = replace(fString, "</font id=beige>", "[/beige]", 1, -1, 1)
		      fString = replace(fString, "<font color=teal>", "[teal]", 1, -1, 1)
		      fString = replace(fString, "</font id=teal>", "[/teal]", 1, -1, 1)
		      fString = replace(fString, "<font color=navy>", "[navy]", 1, -1, 1)
		      fString = replace(fString, "</font id=navy>", "[/navy]", 1, -1, 1)
		      fString = replace(fString, "<font color=maroon>", "[maroon]", 1, -1, 1)
		      fString = replace(fString, "</font id=maroon>", "[/maroon]", 1, -1, 1)
		      fString = replace(fString, "<font color=limegreen>", "[limegreen]", 1, -1, 1)
		      fString = replace(fString, "</font id=limegreen>", "[/limegreen]", 1, -1, 1)

			fString = replace(fString, "<h1>", "[h1]", 1, -1, 1)
			fString = replace(fString, "</h1>", "[/h1]", 1, -1, 1)
			fString = replace(fString, "<h2>", "[h2]", 1, -1, 1)
			fString = replace(fString, "</h2>", "[/h2]", 1, -1, 1)
			fString = replace(fString, "<h3>", "[h3]", 1, -1, 1)
			fString = replace(fString, "</h3>", "[/h3]", 1, -1, 1)
			fString = replace(fString, "<h4>", "[h4]", 1, -1, 1)
			fString = replace(fString, "</h4>", "[/h4]", 1, -1, 1)
			fString = replace(fString, "<h5>", "[h5]", 1, -1, 1)
			fString = replace(fString, "</h5>", "[/h5]", 1, -1, 1)
			fString = replace(fString, "<h6>", "[h6]", 1, -1, 1)
			fString = replace(fString, "</h6>", "[/h6]", 1, -1, 1)
			fString = replace(fString, "<font size=1>", "[size=1]", 1, -1, 1)
			fString = replace(fString, "</font id=size1>", "[/size=1]", 1, -1, 1)
			fString = replace(fString, "<font size=2>", "[size=2]", 1, -1, 1)
			fString = replace(fString, "</font id=size2>", "[/size=2]", 1, -1, 1)
			fString = replace(fString, "<font size=3>", "[size=3]", 1, -1, 1)
			fString = replace(fString, "</font id=size3>", "[/size=3]", 1, -1, 1)
			fString = replace(fString, "<font size=4>", "[size=4]", 1, -1, 1)
			fString = replace(fString, "</font id=size4>", "[/size=4]", 1, -1, 1)
			fString = replace(fString, "<font size=5>", "[size=5]", 1, -1, 1)
			fString = replace(fString, "</font id=size5>", "[/size=5]", 1, -1, 1)
			fString = replace(fString, "<font size=6>", "[size=6]", 1, -1, 1)
			fString = replace(fString, "</font id=size6>", "[/size=6]", 1, -1, 1)
			fString = replace(fString, "<br>","[br]", 1, -1, 1)
			fString = replace(fString, "[/URL]", "[/url]", 1, -1, 1)
			fString = replace(fString, "[URL", "[url", 1, -1, 1)
		    fString = replace(fString, "<div align=left>", "[left]", 1, -1, 1)
		    fString = replace(fString, "</div id=left>", "[/left]", 1, -1, 1)
			fString = replace(fString, "<center>","[center]", 1, -1, 1)
			fString = replace(fString, "</center>","[/center]", 1, -1, 1)
		    fString = replace(fString, "<div align=right>", "[right]", 1, -1, 1)
		    fString = replace(fString, "</div id=right>", "[/right]", 1, -1, 1)
			fString = replace(fString, "<ul>","[list]", 1, -1, 1)
			fString = replace(fString, "</ul>","[/list]", 1, -1, 1)
			fString = replace(fString, "<ol type=1>","[list=1]", 1, -1, 1)
			fString = replace(fString, "</ol id=1>","[/list=1]", 1, -1, 1)
			fString = replace(fString, "<ol type=a>","[list=a]", 1, -1, 1)
			fString = replace(fString, "</ol id=a>","[/list=a]", 1, -1, 1)
			fString = replace(fString, "<li>","[*]", 1, -1, 1)
			fString = replace(fString, "</li>","[/*]", 1, -1, 1)
			fString = replace(fString, "<BLOCKQUOTE id=quote><font size=" & strFooterFontSize & " face=""" & strDefaultFontFace & """ id=quote>引用:<hr height=1 noshade id=quote>","[quote]", 1, -1, 1)
			fString = replace(fString, "<hr height=1 noshade id=quote></BLOCKQUOTE id=quote></font id=quote><font face=""" & strDefaultFontFace & """ size=" & strDefaultFontSize & " id=quote>","[/quote]", 1, -1, 1)

			fString = replace(fString, "<pre id=code><font face=courier size=" & strDefaultFontSize & " id=code>","[code]", 1, -1, 1)
			fString = replace(fString, "</font id=code></pre id=code>","[/code]", 1, -1, 1)
'			fString = replace(fString,"<a href='","[url=", 1, -1, 1)
'			fString = replace(fString,"' target=_blank>", "]", 1, -1, 1)
'			fString = replace(fString,"</a>","[/url]", 1, -1, 1)
		if stricons = "1" then
        'Smile Editing mod By Eric J starts here
        	strsql = "Select smile_url, smile_code from forum_smiles"
            set drs = my_conn.execute(strsql)
            Do until drs.eof
            smile_url = drs("smile_url")
            smile_code = drs("smile_code")
			'fString= replace(fString, "<img src="""&smile_url&" border=0>", "["&smile_code&"]", 1, -1, 1)
			drs.movenext
            loop
            set drs=nothing
        'Ends
		end if

			if strIMGInPosts = "1" then
				fString = replace(fString, "<img src=""","[img]", 1, -1, 1)
				fString = replace(fString, "<img align=right src=""","[img=right]", 1, -1, 1)
				fString = replace(fString, "<img align=left src=""","[img=left]", 1, -1, 1)
				fString = replace(fString, """ border=0>","[/img]", 1, -1, 1)
				fString = replace(fString, """ id=right border=0>","[/img=right]", 1, -1, 1)
				fString = replace(fString, """ id=left border=0>","[/img=left]", 1, -1, 1)
			end if
		end if
	end if
	fString = Replace(fString, "'", "'")
	CleanCode = fString
end function

function Smile(fString)
'Smile Editing Mod by Eric J starts here
        	strsql = "Select smile_url, smile_code from " & strTablePrefix & "smiles"  'Bug
            set drs = my_conn.execute(strsql)
            Do until drs.eof
            smile_url = drs("smile_url")
            smile_code = drs("smile_code")
			
				fString = replace(fString, "["&smile_code&"]", "<img src=""" & strImageURL &smile_url&""" border=0>")
			drs.movenext
            loop
            set drs=nothing
	Smile = fString
'Ends here
end function

Function ChkBadWords(fString)
  	Dim regEx, str1,chrLastLetter               ' Create variables.
  	str1 = fString
	bwords = split(strBadWords, "|")
	for i = 0 to ubound(bwords)
  		Set regEx = New RegExp            ' Create regular expression.
  		chrLastLetter = right(bwords(i),1)
		regEx.Pattern = "\b" & left(bwords(i),len(bwords(i))-1) & "(" & chrLastLetter & "|" & chrLastLetter & "s|" & chrLastLetter & "es)" & "\b"            ' Set pattern.
  		regEx.IgnoreCase = FALSE            ' Make case insensitive.
  		str1 = regEx.Replace(fstring, string(len(regEx.Pattern),"*"))
		do until str1 = fstring
			fstring = str1
			str1 = regEx.Replace(fstring, string(len(bwords(i)),"*"))
		loop
		
 	next
	ChkBadWords = str1
	
End Function



function HTMLEncode(fString)

	fString = replace(fString, ">", "&gt;")
	fString = replace(fString, "<", "&lt;")

	HTMLEncode = fString
end function

function HTMLDecode(fString)

	fString = replace(fString, "&gt;", ">")
	fString = replace(fString, "&lt;", "<")

	HTMLDecode = fString
end function


function ChkString(fString,fField_Type) '## Types - name, password, title, message, url, urlpath, email, number, list
	fString = trim(fString)
	if fString = "" then
		fString = " "
	end if
	if fField_Type = "title" or fField_Type = "preview" then
			fString = Replace(fString, "'", "''")
			if strAllowHTML <> "1" then
						fString = HTMLEncode(fString)
			end if
			ChkBadWords(fString)
			ChkString = fString
			exit function
	end if
	if fField_Type = "decode" then
			fString = HTMLDecode(fString)
			ChkString = fString
			exit function
	end if
	if fField_Type = "urlpath" then
			fString = Server.URLEncode(fString)
			ChkString = fString
			exit function
	end if
	if fField_Type = "SQLString" then
		fString = Replace(fString, "'", "''")
		fString = HTMLEncode(fString)
		ChkString = fString
		exit function
	end if
	if fField_Type = "JSurlpath" then
		fString = Replace(fString, "'", "\'")
		fString = Server.URLEncode(fString)
		ChkString = fString
		exit function
	end if
	if fField_Type = "display" then
		if strAllowHTML <> "1" then
			fString = HTMLEncode(fString)
		end if
		ChkString = fString
		exit function
	elseif fField_Type = "message" then
			if strAllowHTML <> "1" then
				fString = HTMLEncode(fString)
			end if
	elseif fField_Type = "preview" then
			if strAllowHTML <> "1" then
				fString = HTMLEncode(fString)
			end if
	elseif fField_Type = "hidden" then
		fString = HTMLEncode(fString)
	end if
	if strAllowForumCode = "1" and fField_Type <> "signature" then
		fString = doCode(fString, "[b]", "[/b]", "<b>", "</b>")
		fString = doCode(fString, "[s]", "[/s]", "<s>", "</s>")
		fString = doCode(fString, "[strike]", "[/strike]", "<s>", "</s>")
		fString = doCode(fString, "[u]", "[/u]", "<u>", "</u>")
		fString = doCode(fString, "[i]", "[/i]", "<i>", "</i>")
		if fField_Type <> "title" then
			fString = doCode(fString, "[font=Andale Mono]", "[/font=Andale Mono]", "<font face='Andale Mono'>", "</font id='Andale Mono'>")
			fString = doCode(fString, "[font=Arial]", "[/font=Arial]", "<font face='Arial'>", "</font id='Arial'>")
			fString = doCode(fString, "[font=Arial Black]", "[/font=Arial Black]", "<font face='Arial Black'>", "</font id='Arial Black'>")
			fString = doCode(fString, "[font=Book Antiqua]", "[/font=Book Antiqua]", "<font face='Book Antiqua'>", "</font id='Book Antiqua'>")
			fString = doCode(fString, "[font=Century Gothic]", "[/font=Century Gothic]", "<font face='Century Gothic'>", "</font id='Century Gothic'>")
			fString = doCode(fString, "[font=Courier New]", "[/font=Courier New]", "<font face='Courier New'>", "</font id='Courier New'>")
			fString = doCode(fString, "[font=Comic Sans MS]", "[/font=Comic Sans MS]", "<font face='Comic Sans MS'>", "</font id='Comic Sans MS'>")
			fString = doCode(fString, "[font=Georgia]", "[/font=Georgia]", "<font face='Georgia'>", "</font id='Georgia'>")
			fString = doCode(fString, "[font=Impact]", "[/font=Impact]", "<font face='Impact'>", "</font id='Impact'>")
			fString = doCode(fString, "[font=Tahoma]", "[/font=Tahoma]", "<font face='Tahoma'>", "</font id='Tahoma'>")
			fString = doCode(fString, "[font=Times New Roman]", "[/font=Times New Roman]", "<font face='Times New Roman'>", "</font id='Times New Roman'>")
			fString = doCode(fString, "[font=Trebuchet MS]", "[/font=Trebuchet MS]", "<font face='Trebuchet MS'>", "</font id='Trebuchet MS'>")
			fString = doCode(fString, "[font=Script MT Bold]", "[/font=Script MT Bold]", "<font face='Script MT Bold'>", "</font id='Script MT Bold'>")
			fString = doCode(fString, "[font=Stencil]", "[/font=Stencil]", "<font face='Stencil'>", "</font id='Stencil'>")
			fString = doCode(fString, "[font=宋体]", "[/font=宋体]", "<font face='宋体'>", "</font id='宋体'>")
			fString = doCode(fString, "[font=Lucida Console]", "[/font=Lucida Console]", "<font face='Lucida Console'>", "</font id='Lucida Console'>")

			fString = doCode(fString, "[red]", "[/red]", "<font color=red>", "</font id=red>")
			fString = doCode(fString, "[green]", "[/green]", "<font color=green>", "</font id=green>")
			fString = doCode(fString, "[blue]", "[/blue]", "<font color=blue>", "</font id=blue>")
			fString = doCode(fString, "[white]", "[/white]", "<font color=white>", "</font id=white>")
			fString = doCode(fString, "[purple]", "[/purple]", "<font color=purple>", "</font id=purple>")
			fString = doCode(fString, "[yellow]", "[/yellow]", "<font color=yellow>", "</font id=yellow>")
			fString = doCode(fString, "[violet]", "[/violet]", "<font color=violet>", "</font id=violet>")
			fString = doCode(fString, "[brown]", "[/brown]", "<font color=brown>", "</font id=brown>")
			fString = doCode(fString, "[black]", "[/black]", "<font color=black>", "</font id=black>")
			fString = doCode(fString, "[pink]", "[/pink]", "<font color=pink>", "</font id=pink>")
			fString = doCode(fString, "[orange]", "[/orange]", "<font color=orange>", "</font id=orange>")
			fString = doCode(fString, "[gold]", "[/gold]", "<font color=gold>", "</font id=gold>")

			fString = doCode(fString, "[beige]", "[/beige]", "<font color=beige>", "</font id=beige>")
			fString = doCode(fString, "[teal]", "[/teal]", "<font color=teal>", "</font id=teal>")
			fString = doCode(fString, "[navy]", "[/navy]", "<font color=navy>", "</font id=navy>")
			fString = doCode(fString, "[maroon]", "[/maroon]", "<font color=maroon>", "</font id=maroon>")
			fString = doCode(fString, "[limegreen]", "[/limegreen]", "<font color=limegreen>", "</font id=limegreen>")

			fString = doCode(fString, "[h1]", "[/h1]", "<h1>", "</h1>")
			fString = doCode(fString, "[h2]", "[/h2]", "<h2>", "</h2>")
			fString = doCode(fString, "[h3]", "[/h3]", "<h3>", "</h3>")
			fString = doCode(fString, "[h4]", "[/h4]", "<h4>", "</h4>")
			fString = doCode(fString, "[h5]", "[/h5]", "<h5>", "</h5>")
			fString = doCode(fString, "[h6]", "[/h6]", "<h6>", "</h6>")
			fString = doCode(fString, "[size=1]", "[/size=1]", "<font size=1>", "</font id=size1>")
			fString = doCode(fString, "[size=2]", "[/size=2]", "<font size=2>", "</font id=size2>")
			fString = doCode(fString, "[size=3]", "[/size=3]", "<font size=3>", "</font id=size3>")
			fString = doCode(fString, "[size=4]", "[/size=4]", "<font size=4>", "</font id=size4>")
			fString = doCode(fString, "[size=5]", "[/size=5]", "<font size=5>", "</font id=size5>")
			fString = doCode(fString, "[size=6]", "[/size=6]", "<font size=6>", "</font id=size6>")
			fString = doCode(fString, "[list]", "[/list]", "<ul>", "</ul>")
			fString = doCode(fString, "[list=1]", "[/list=1]", "<ol type=1>", "</ol id=1>")
			fString = doCode(fString, "[list=a]", "[/list=a]", "<ol type=a>", "</ol id=a>")
			fString = doCode(fString, "[*]", "[/*]", "<li>", "</li>")
			fString = doCode(fString, "[left]", "[/left]", "<div align=left>", "</div id=left>")
			fString = doCode(fString, "[center]", "[/center]", "<center>", "</center>")
			fString = doCode(fString, "[centre]", "[/centre]", "<center>", "</center>")
			fString = doCode(fString, "[right]", "[/right]", "<div align=right>", "</div id=right>")
			fString = doCode(fString, "[code]", "[/code]", "<pre id=code><font face=courier size=" & strDefaultFontSize & " id=code>", "</font id=code></pre id=code>")
			fString = doCode(fString, "[quote]", "[/quote]", "<BLOCKQUOTE id=quote><font size=" & strFooterFontSize & " face=""" & strDefaultFontFace & """ id=quote>quote:<hr height=1 noshade id=quote>", "<hr height=1 noshade id=quote></font id=quote></BLOCKQUOTE id=quote>")
			fString = doCode(fString, "[url=&quot;", "&quot;]", "[url=""", """]")
			fString = doCode(fString, "[URL=&quot;", "&quot;]", "[url=""", """]")
'			fString = doCode(fString, "[url", "[/url]", "<a>", "</a>")
			fString = replace(fString, "[br]", "<br>", 1, -1, 1)
	if stricons = "1" and _
	fField_Type <> "title" and _
		fField_Type <> "hidden" then
		fString= smile(fString)
	end if
			
			if strIMGInPosts = "1" then
				fString = doCode(fString, "[img]","[/img]","<img src=""",""" border=0>")
				fString = doCode(fString, "[image]","[/image]","<img src=""",""" border=0>")
				fString = doCode(fString, "[img=right]","[/img=right]","<img align=right src=""",""" id=right border=0>")
				fString = doCode(fString, "[image=right]","[/image=right]","<img align=right src=""",""" id=right border=0>")
				fString = doCode(fString, "[img=left]","[/img=left]","<img align=left src=""",""" id=left border=0>")
				fString = doCode(fString, "[image=left]","[/image=left]","<img align=left src=""",""" id=left border=0>")
			end if
		end if
	end if
	if strIcons = "1" and _
	fField_Type <> "title" and _
	fField_Type <> "hidden" then
		fString= smile(fString)
	end if
	if fField_Type <> "hidden" and _
	fField_Type <> "preview" then
		fString = Replace(fString, "'", "''")
	end if
	ChkString = fString
end function

function ChkDateTime(fDateTime)
	if fDateTime = "" then
		exit function
	end if
	if IsDate(fDateTime) then
		select case strDateType
			case "dmy"
				ChkDateTime = Mid(fDateTime,7,2) & "/" & _
				Mid(fDateTime,5,2) & "/" & _
				Mid(fDateTime,1,4)
			case "mdy"
				ChkDateTime = Mid(fDateTime,5,2) & "/" & _
				Mid(fDateTime,7,2) & "/" & _
				Mid(fDateTime,1,4)
			case "ymd"
				ChkDateTime = Mid(fDateTime,1,4) & "/" & _
				Mid(fDateTime,5,2) & "/" & _
				Mid(fDateTime,7,2)
			case "ydm"
				ChkDateTime =Mid(fDateTime,1,4) & "/" & _
				Mid(fDateTime,7,2) & "/" & _
				Mid(fDateTime,5,2)
			case "dmmy"
				ChkDateTime = Mid(fDateTime,7,2) & " " & _
				Monthname(Mid(fDateTime,5,2),1) & " " & _
				Mid(fDateTime,1,4)
			case "mmdy"
				ChkDateTime = Monthname(Mid(fDateTime,5,2),1) & " " & _
				Mid(fDateTime,7,2) & " " & _
				Mid(fDateTime,1,4)
			case "ymmd"
				ChkDateTime = Mid(fDateTime,1,4) & " " & _
				Monthname(Mid(fDateTime,5,2),1) & " " & _
				Mid(fDateTime,7,2)
			case "ydmm"
				ChkDateTime = Mid(fDateTime,1,4) & " " & _
				Mid(fDateTime,7,2) & " " & _
				Monthname(Mid(fDateTime,5,2),1)
			case "dmmmy"
				ChkDateTime = Mid(fDateTime,7,2) & " " & _
				Monthname(Mid(fDateTime,5,2),0) & " " & _
				Mid(fDateTime,1,4)
			case "mmmdy"
				ChkDateTime = Monthname(Mid(fDateTime,5,2),0) & " " & _
				Mid(fDateTime,7,2) & " " & _
				Mid(fDateTime,1,4)
			case "ymmmd"
				ChkDateTime = Mid(fDateTime,1,4) & " " & _
				Monthname(Mid(fDateTime,5,2),0) & " " & _
				Mid(fDateTime,7,2)
			case "ydmmm"
				ChkDateTime = Mid(fDateTime,1,4) & " " & _
				Mid(fDateTime,7,2) & " " & _
				Monthname(Mid(fDateTime,5,2),0)
			case else
				ChkDateTime = doublenum(Mid(fDateTime,5,2)) & "/" & _
				Mid(fDateTime,7,2) & "/" & _
				Mid(fDateTime,1,4)
		end select
		if strTimeType = 12 then
			if cint(Mid(fDateTime, 9,2)) > 12 then
				ChkDateTime = ChkDateTime & " " & _
				(cint(Mid(fDateTime, 9,2)) -12) & ":" & _
				Mid(fDateTime, 11,2) & ":" & _
				Mid(fDateTime, 13,2) & " " & "PM"
			elseif cint(Mid(fDateTime, 9,2)) = 12 then
				ChkDateTime = ChkDateTime & " " & _
				cint(Mid(fDateTime, 9,2)) & ":" & _
				Mid(fDateTime, 11,2) & ":" & _
				Mid(fDateTime, 13,2) & " " & "PM"
			elseif cint(Mid(fDateTime, 9,2)) = 0 then
				ChkDateTime = ChkDateTime & " " & _
				(cint(Mid(fDateTime, 9,2)) +12) & ":" & _
				Mid(fDateTime, 11,2) & ":" & _
				Mid(fDateTime, 13,2) & " " & "AM"
			else
				ChkDateTime = ChkDateTime & " " & _
				Mid(fDateTime, 9,2) & ":" & _
				Mid(fDateTime, 11,2) & ":" & _
				Mid(fDateTime, 13,2) & " " & "AM"
			end if
		else
			ChkDateTime = ChkDateTime & " " & _
			Mid(fDateTime, 9,2) & ":" & _
			Mid(fDateTime, 11,2) & ":" & _
			Mid(fDateTime, 13,2) 
		end if
	end if
end function

function ChkDateFormat(strDateTime)
	ChkDateFormat =  isdate("" & Mid(strDateTime, 5,2) & "/" & Mid(strDateTime, 7,2) & "/" & Mid(strDateTime, 1,4) & " " & Mid(strDateTime, 9,2) & ":" & Mid(strDateTime, 11,2) & ":" & Mid(strDateTime, 13,2) & "")
end function

function StrToDate(strDateTime)
	if ChkDateFormat(strDateTime) then
		StrToDate = cdate("" & Mid(strDateTime, 5,2) & "/" & Mid(strDateTime, 7,2) & "/" & Mid(strDateTime, 1,4) & " " & Mid(strDateTime, 9,2) & ":" & Mid(strDateTime, 11,2) & ":" & Mid(strDateTime, 13,2) & "")
	else
		StrToDate = "" & strForumTimeAdjust
	end if
end function

function DateToStr(dtDateTime)
	DateToStr = year(dtDateTime) & doublenum(Month(dtdateTime)) & doublenum(Day(dtdateTime)) & doublenum(Hour(dtdateTime)) & doublenum(Minute(dtdateTime)) & doublenum(Second(dtdateTime)) & ""
end function

function ReadLastHereDate(UserName)
	dim TempLastHereDate
	dim rs_date

	'## Forum_SQL
	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.M_LASTHEREDATE "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS."&Strdbntsqlname&" = '" & UserName & "' "

	
	set rs_date = my_conn.Execute (strSql)

	if (rs_date.BOF and rs_date.EOF) then
		TempLastHereDate = DateAdd("d",-10,strForumTimeAdjust)
	else
		TempLastHereDate = StrToDate(rs_date("M_LASTHEREDATE"))
		if TempLastHereDate = "" or IsNull(TempLastHereDate) then
			TempLastHereDate = DateAdd("d",-10,strForumTimeAdjust)
		end if	
	end if

	rs_date.close	
	set rs_date = nothing

	'## Forum_SQL - Do DB Update
	strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " SET M_LASTHEREDATE = '" & DateToStr(strForumTimeAdjust) & "'"
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS." & strDBNTSQLName & " = '" & UserName & "' "

	
	my_conn.Execute (strSql)	

	ReadLastHereDate = DateToStr(TempLastHereDate)
end function

function getMemberID(fUser_Name)

	'## Forum_SQL
	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE M_NAME = '" & fUser_Name & "'"

	rsGetMemberID = my_Conn.Execute(strSql)
	
	getMemberID = rsGetMemberID("MEMBER_ID")

end function

function ChkDate(fDate)
	if fDate = "" or vartype(fDate) = vbNull then
		exit function
	end if
	'if IsDate(fDate) then
		select case strDateType
			case "dmy"
				ChkDate = Mid(fDate,7,2) & "/" & _
				Mid(fDate,5,2) & "/" & _
				Mid(fDate,1,4)
			case "mdy"
				ChkDate = Mid(fDate,5,2) & "/" & _
				Mid(fDate,7,2) & "/" & _
				Mid(fDate,1,4)
			case "ymd"
				ChkDate = Mid(fDate,1,4) & "/" & _
				Mid(fDate,5,2) & "/" & _
				Mid(fDate,7,2)
			case "ydm"
				ChkDate =Mid(fDate,1,4) & "/" & _
				Mid(fDate,7,2) & "/" & _
				Mid(fDate,5,2)
			case "dmmy"
				ChkDate = Mid(fDate,7,2) & " " & _
				Monthname(Mid(fDate,5,2),1) & " " & _
				Mid(fDate,1,4)
			case "mmdy"
				ChkDate = Monthname(Mid(fDate,5,2),1) & " " & _
				Mid(fDate,7,2) & " " & _
				Mid(fDate,1,4)
			case "ymmd"
				ChkDate = Mid(fDate,1,4) & " " & _
				Monthname(Mid(fDate,5,2),1) & " " & _
				Mid(fDate,7,2)
			case "ydmm"
				ChkDate = Mid(fDate,1,4) & " " & _
				Mid(fDate,7,2) & " " & _
				Monthname(Mid(fDate,5,2),1)
			case "dmmmy"
				ChkDate = Mid(fDate,7,2) & " " & _
				Monthname(Mid(fDate,5,2),0) & " " & _
				Mid(fDate,1,4)
			case "mmmdy"
				ChkDate = Monthname(Mid(fDate,5,2),0) & " " & _
				Mid(fDate,7,2) & " " & _
				Mid(fDate,1,4)
			case "ymmmd"
				ChkDate = Mid(fDate,1,4) & " " & _
				Monthname(Mid(fDate,5,2),0) & " " & _
				Mid(fDate,7,2)
			case "ydmmm"
				ChkDate = Mid(fDate,1,4) & " " & _
				Mid(fDate,7,2) & " " & _
				Monthname(Mid(fDate,5,2),0)
			case else
				ChkDate = Mid(fDate,5,2) & "/" & _
				Mid(fDate,7,2) & "/" & _
				Mid(fDate,1,4)
		End Select
	'end if
end function

function ChkTime(fTime)
	if fTime = "" then
		exit function
	end if

	if strTimeType = 12 then
		if cint(Mid(fTime, 9,2)) > 12 then
			ChkTime = ChkTime & " " & _
			(cint(Mid(fTime, 9,2)) -12) & ":" & _
			Mid(fTime, 11,2) & ":" & _
			Mid(fTime, 13,2) & " " & "PM"
		elseif cint(Mid(fTime, 9,2)) = 12 then
			ChkTime = ChkTime & " " & _
			cint(Mid(fTime, 9,2)) & ":" & _
			Mid(fTime, 11,2) & ":" & _
			Mid(fTime, 13,2) & " " & "PM"
		elseif cint(Mid(fTime, 9,2)) = 0 then
			ChkTime = ChkTime & " " & _
			(cint(Mid(fTime, 9,2)) +12) & ":" & _
			Mid(fTime, 11,2) & ":" & _
			Mid(fTime, 13,2) & " " & "AM"
		else
			ChkTime = ChkTime & " " & _
			Mid(fTime, 9,2) & ":" & _
			Mid(fTime, 11,2) & ":" & _
			Mid(fTime, 13,2) & " " & "AM"
		end if
		
	else
		ChkTime = ChkTime & " " & _
		Mid(fTime, 9,2) & ":" & _
		Mid(fTime, 11,2) & ":" & _
		Mid(fTime, 13,2) 
	end if
end function

function EmailField(fTestString) 
	TheAt = Instr(2, fTestString, "@")
	if TheAt = 0 then 
		EmailField = 0
	else
		TheDot = Instr(cint(TheAt) + 2, fTestString, ".")
		if TheDot = 0 then
			EmailField = 0
		else
			if cint(TheDot) + 1 > Len(fTestString) then
				EmailField = 0
			else
				EmailField = -1
			end if
		end if
	end if
end function 

function ChkIsNew(fDateTime,fType)
	if fType = "Topic" then
		if strHotTopic = "1" then
			if fDateTime > Session(strCookieURL & "last_here_date") then
				if rs("T_REPLIES") >= intHotTopicNum then
				        ChkIsNew =  "<img src=" & strImageURL & "icon_folder_new_hot.gif height=15 width=15 alt=热门主题 border=0>"
				else
				        ChkIsNew =  "<img src=" & strImageURL & "icon_folder_new.gif height=15 width=15 alt=新主题 border=0>"
				end if
			else
				if rs("T_REPLIES") >= intHotTopicNum then
				        ChkIsNew =  "<img src=" & strImageURL & "icon_folder_hot.gif height=15 width=15 alt=热门主题 border=0>"
				else
				        ChkIsNew = "<img src=" & strImageURL & "icon_folder.gif height=15 width=15 border=0>" 
				end if
			end if
		else
			if fDateTime > Session(strCookieURL & "last_here_date") then
				ChkIsNew =  "<img src=" & strImageURL & "icon_folder_new.gif height=15 width=15 alt=新主题 border=0>" 
			else
				ChkIsNew = "<img src=" & strImageURL & "icon_folder.gif height=15 width=15 border=0>" 
			end if
		end if
	else
		if rsForum("F_STATUS") <> 0 then
			if fDateTime > Session(strCookieURL & "last_here_date") then
				ChkIsNew =  "<IMG src='" & strImageURL & "icon_folder_new.gif' height=15 width=15 border=0 hspace=0 alt='New Posts'>" 
			Else
				ChkIsNew = "<IMG src='" & strImageURL & "icon_folder.gif' height=15 width=15 border=0 hspace=0 alt='Old Posts'>" 
			end if
		else
			ChkIsNew = "<IMG src='" & strImageURL & "icon_folder_locked.gif' height=15 width=15 border=0 hspace=0 alt='论坛已被锁定'>"
		end if
	end if
end function

function ChkQuoteOk(fString)

	ChkQuoteOk = not(InStr(1, fString, "'", 0) > 0)

end function

function ChkUser(fName, fPassword)

	'## Forum_SQL
	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_LEVEL, " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.M_PASSWORD "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS." & strDBNTSQLName & " = '" & fName & "' "
	if strAuthType="db" then
		strSql = strSql & " AND   " & strMemberTablePrefix & "MEMBERS.M_PASSWORD = '" & fPassword &"'"
	End IF
	strSql = strSql & " AND   " & strMemberTablePrefix & "MEMBERS.M_STATUS = " & 1

	set rsCheck = my_Conn.Execute (strSql)

	if rsCheck.BOF or rsCheck.EOF or not(ChkQuoteOk(fName)) or not(ChkQuoteOk(fPassword)) then
		ChkUser = 0
	else
		if cstr(rsCheck("MEMBER_ID")) = Request.Form("Author") then
			ChkUser = 1 '## Author
		else
			Select case cint(rsCheck("M_LEVEL"))
				case 1
					ChkUser = 2 '## Normal User
				case 2
					ChkUser = 3 '## Moderator
				case 3
					ChkUser = 4 '## Admin
				case else
					ChkUser = cint(rsCheck("M_LEVEL"))
			End Select
		end if	
	end if

	rsCheck.close
	set rsCheck = nothing

end function

function ChkUser2(fName, fPassword)

	'## Forum_SQL
	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_LEVEL, " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.M_PASSWORD "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	StrSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS." & strDBNTSQLName & " = '" & fName & "' "
	if strAuthType="db" then	
		strSql = strSql & " AND   " & strMemberTablePrefix & "MEMBERS.M_PASSWORD = '" & fPassword &"'"
	End If
	strSql = strSql & " AND   " & strMemberTablePrefix & "MEMBERS.M_STATUS = " & 1

	on error resume next
	set rsCheck = my_Conn.Execute (strSql)
	
	for counter = 0 to my_Conn.Errors.Count -1
		if my_Conn.Errors(counter).Number <> 0 or Err.number > 0 then 
			ChkUser2 = -1
			my_Conn.Errors.Clear 
		end if
	next

	if rsCheck.BOF or rsCheck.EOF or not(ChkQuoteOk(fName)) or not(ChkQuoteOk(fPassword)) or ChkUser2 = -1 then
		ChkUser2 = 0 '##  Invalid Password
	else
		if cint(rsCheck("MEMBER_ID")) = cint(Request.QueryString("Author")) then 
			ChkUser2 = 1 '## Author
		else
			select case cint(rsCheck("M_LEVEL"))
				case 1
					ChkUser2 = 2 '## Normal User
				case 2
					ChkUser2 = 3 '## Moderator
				case 3
					ChkUser2 = 4 '## Admin
				case else
					ChkUser2 = cint(rsCheck("M_LEVEL"))
			end select
		end if	
	end if

	rsCheck.close	
	set rsCheck = nothing

end function

function ChkUser3(fName, fPassword, fReply)

	'## Forum_SQL
	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_LEVEL, " & strMemberTablePrefix & "MEMBERS.M_NAME, " & strMemberTablePrefix & "MEMBERS.M_PASSWORD, " & strTablePrefix & "REPLY.R_AUTHOR "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS, " & strTablePrefix & "REPLY "
	StrSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS." & strDBNTSQLName & " = '" & fName & "' "
	if strAuthType="db" then	
		strSql = strSql & " AND   " & strMemberTablePrefix & "MEMBERS.M_PASSWORD = '" & fPassword &"' "
	End If
	strSql = strSql & " AND   " & strTablePrefix & "REPLY.REPLY_ID = " & fReply
	strSql = strSql & " AND   " & strMemberTablePrefix & "MEMBERS.M_STATUS = " & 1

	set rsCheck = my_Conn.Execute (strSql)

	if rsCheck.BOF or rsCheck.EOF or not(ChkQuoteOk(fName)) or not(ChkQuoteOk(fPassword)) then
		ChkUser3 = 0 '##  Invalid Password
	else
		if cint(rsCheck("MEMBER_ID")) = cint(rsCheck("R_AUTHOR")) then 
			ChkUser3 = 1 '## Author
		else
			Select case cint(rsCheck("M_LEVEL"))
				case 1
					ChkUser3 = 2 '## Normal User
				case 2
					ChkUser3 = 3 '## Moderator
				case 3
					ChkUser3 = 4 '## Admin
				case else
					ChkUser3 = cint(rsCheck("M_LEVEL"))
			End Select
		end if	
	end if

	rsCheck.close	
	set rsCheck = nothing

end function

function GetSig(fUser_Name)

	'## Forum_SQL
	strSql = "SELECT M_SIG "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE M_NAME = '" & Request.Form("UserName") & "'"

	set rsSig = my_Conn.Execute (strSql)

	if rsSig.EOF or rsSig.BOF then
		'## Do Nothing
	else
		GetSig = rsSig("M_SIG")
	end if

	rsSig.close
	set rsSig = nothing

end function

function DoDropDown(fTableName, fDisplayField, fValueField, fSelectValue, fName)


	'## Forum_SQL
	strSql = "SELECT " & fDisplayField & ", " & fValueField 
	strSql = strSql & " FROM " & fTableName

	rsdrop.Open strSql, my_Conn
		
	Response.Write "<Select Name='" & fName & "'>"
	if rsdrop.EOF or rsdrop.BOF then 
		Response.Write "<Option>No Items Found</option>"  & vbCrLf
	else
		do until rsdrop.EOF
			if rs(fValueField) = cint(fSelectValue) then
				Response.Write "<option value='" & rsdrop(fValueField) & "' Selected>"
				Response.Write rsdrop(fDisplayField) & "</option>" & vbCrLf
			else
				Response.Write "<option value='" & rsdrop(fValueField) & "'>"
				Response.Write rsdrop(fDisplayField) & "</option>" & vbCrLf
			end if
			rsdrop.MoveNext
		loop
	end if
	Response.Write "</select>" & vbCrLf

	rsdrop.Close
	set rsdrop = nothing	

end function

sub DoULastPost(sUser_Name)

	'## Forum_SQL - Updates the M_LASTPOSTDATE in the FORUM_MEMBERS table
	strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " SET M_LASTPOSTDATE = '" & DateToStr(strForumTimeAdjust) & "' "
	strSql = strSql & " WHERE " & strDBNTSQLName & " = '" & sUser_Name & "'"
	my_Conn.Execute (strSql)
	
end sub

'##############################################
'##            Ranks and Stars               ##
'##############################################

function getMember_Level(fM_TITLE, fM_LEVEL, fM_POSTS)

dim Member_Level

	Member_Level = ""
	if Trim(fM_TITLE) <> "" then
		Member_Level = fM_TITLE
	else
		select case fM_LEVEL
			case "1"  
				if (fM_POSTS < intRankLevel1) then Member_Level = Member_Level & strRankLevel0
				if (fM_POSTS >= intRankLevel1) and (fM_POSTS < intRankLevel2) then Member_Level = Member_Level & strRankLevel1
				if (fM_POSTS >= intRankLevel2) and (fM_POSTS < intRankLevel3) then Member_Level = Member_Level & strRankLevel2
				if (fM_POSTS >= intRankLevel3) and (fM_POSTS < intRankLevel4) then Member_Level = Member_Level & strRankLevel3
				if (fM_POSTS >= intRankLevel4) and (fM_POSTS < intRankLevel5) then Member_Level = Member_Level & strRankLevel4
				if (fM_POSTS >= intRankLevel5) then Member_Level = Member_Level & strRankLevel5
			case "2"
				Member_Level = Member_Level & strRankMod
			case "3"
				Member_Level = Member_Level & strRankAdmin
			case else  
				Member_Level = Member_Level & "Error" 
		end select
	end if
	
getMember_Level = Member_Level
	
end function

function getStar_Level(fM_LEVEL, fM_POSTS)

dim Star_Level

	Star_Level = ""
	select case fM_LEVEL
		case "1"
			if (fM_POSTS < intRankLevel1) then Star_Level = Star_Level & ""
			if (fM_POSTS >= intRankLevel1) and (fM_POSTS < intRankLevel2) then Star_Level = Star_Level & "<img src=" & strImageURL & "icon_star_" & strRankColor1 & ".gif border=0>"
			if (fM_POSTS >= intRankLevel2) and (fM_POSTS < intRankLevel3) then Star_Level = Star_Level & "<img src=" & strImageURL & "icon_star_" & strRankColor2 & ".gif border=0><img src=" & strImageURL & "icon_star_" & strRankColor2 & ".gif border=0>"
			if (fM_POSTS >= intRankLevel3) and (fM_POSTS < intRankLevel4) then Star_Level = Star_Level & "<img src=" & strImageURL & "icon_star_" & strRankColor3 & ".gif border=0><img src=" & strImageURL & "icon_star_" & strRankColor3 & ".gif border=0><img src=" & strImageURL & "icon_star_" & strRankColor3 & ".gif border=0>"
			if (fM_POSTS >= intRankLevel4) and (fM_POSTS < intRankLevel5) then Star_Level = Star_Level & "<img src=" & strImageURL & "icon_star_" & strRankColor4 & ".gif border=0><img src=" & strImageURL & "icon_star_" & strRankColor4 & ".gif border=0><img src=" & strImageURL & "icon_star_" & strRankColor4 & ".gif border=0><img src=" & strImageURL & "icon_star_" & strRankColor4 & ".gif border=0>"
			if (fM_POSTS >= intRankLevel5) then Star_Level = Star_Level & "<img src=" & strImageURL & "icon_star_" & strRankColor5 & ".gif border=0><img src=" & strImageURL & "icon_star_" & strRankColor5 & ".gif border=0><img src=" & strImageURL & "icon_star_" & strRankColor5 & ".gif border=0><img src=" & strImageURL & "icon_star_" & strRankColor5 & ".gif border=0><img src=" & strImageURL & "icon_star_" & strRankColor5 & ".gif border=0>"
		case "2" 
			if fM_POSTS < intRankLevel1 then Star_Level = Star_Level & ""
			if (fM_POSTS >= intRankLevel1) and (fM_POSTS < intRankLevel2) then Star_Level = Star_Level & "<img src=" & strImageURL & "icon_star_" & strRankColorMod & ".gif border=0>"
			if (fM_POSTS >= intRankLevel2) and (fM_POSTS < intRankLevel3) then Star_Level = Star_Level & "<img src=" & strImageURL & "icon_star_" & strRankColorMod & ".gif border=0><img src=" & strImageURL & "icon_star_" & strRankColorMod & ".gif border=0>"
			if (fM_POSTS >= intRankLevel3) and (fM_POSTS < intRankLevel4) then Star_Level = Star_Level & "<img src=" & strImageURL & "icon_star_" & strRankColorMod & ".gif border=0><img src=" & strImageURL & "icon_star_" & strRankColorMod & ".gif border=0><img src=" & strImageURL & "icon_star_" & strRankColorMod & ".gif border=0>"
			if (fM_POSTS >= intRankLevel4) and (fM_POSTS < intRankLevel5) then Star_Level = Star_Level & "<img src=" & strImageURL & "icon_star_" & strRankColorMod & ".gif border=0><img src=" & strImageURL & "icon_star_" & strRankColorMod & ".gif border=0><img src=" & strImageURL & "icon_star_" & strRankColorMod & ".gif border=0><img src=" & strImageURL & "icon_star_" & strRankColorMod & ".gif border=0>"
			if (fM_POSTS >= intRankLevel5) then Star_Level = Star_Level & "<img src=" & strImageURL & "icon_star_" & strRankColorMod & ".gif border=0><img src=" & strImageURL & "icon_star_" & strRankColorMod & ".gif border=0><img src=" & strImageURL & "icon_star_" & strRankColorMod & ".gif border=0><img src=" & strImageURL & "icon_star_" & strRankColorMod & ".gif border=0><img src=" & strImageURL & "icon_star_" & strRankColorMod & ".gif border=0>"
		case "3"
			Star_Level = "" 
			if (fM_POSTS < intRankLevel1) then Star_Level = Star_Level & ""
			if (fM_POSTS >= intRankLevel1) and (fM_POSTS < intRankLevel2) then Star_Level = Star_Level & "<img src=" & strImageURL & "icon_star_" & strRankColorAdmin & ".gif border=0>"
			if (fM_POSTS >= intRankLevel2) and (fM_POSTS < intRankLevel3) then Star_Level = Star_Level & "<img src=" & strImageURL & "icon_star_" & strRankColorAdmin & ".gif border=0><img src=" & strImageURL & "icon_star_" & strRankColorAdmin & ".gif border=0>"
			if (fM_POSTS >= intRankLevel3) and (fM_POSTS < intRankLevel4) then Star_Level = Star_Level & "<img src=" & strImageURL & "icon_star_" & strRankColorAdmin & ".gif border=0><img src=" & strImageURL & "icon_star_" & strRankColorAdmin & ".gif border=0><img src=" & strImageURL & "icon_star_" & strRankColorAdmin & ".gif border=0>"
			if (fM_POSTS >= intRankLevel4) and (fM_POSTS < intRankLevel5) then Star_Level = Star_Level & "<img src=" & strImageURL & "icon_star_" & strRankColorAdmin & ".gif border=0><img src=" & strImageURL & "icon_star_" & strRankColorAdmin & ".gif border=0><img src=" & strImageURL & "icon_star_" & strRankColorAdmin & ".gif border=0><img src=" & strImageURL & "icon_star_" & strRankColorAdmin & ".gif border=0>"
			if (fM_POSTS >= intRankLevel5) then Star_Level = Star_Level & 								 "<img src=" & strImageURL & "icon_star_" & strRankColorAdmin & ".gif border=0><img src=" & strImageURL & "icon_star_" & strRankColorAdmin & ".gif border=0><img src=" & strImageURL & "icon_star_" & strRankColorAdmin & ".gif border=0><img src=" & strImageURL & "icon_star_" & strRankColorAdmin & ".gif border=0><img src=" & strImageURL & "icon_star_" & strRankColorAdmin & ".gif border=0>"
		case else  
			Star_Level = Star_Level & "Error"
	end select

getStar_Level = Star_Level


end function

'##############################################
'##           Multi-Moderators               ##
'##############################################

function chkForumModerator(fForum_ID, fMember_Name)

	'## Forum_SQL 
	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS." & strDBNTSQLName & " = '" & fMember_Name & "'"
	

	set rsUsrName = my_Conn.Execute (strSql)

	if rsUsrName.EOF or rsUsrName.BOF or not(ChkQuoteOk(fMember_Name)) or not(ChkQuoteOk(fForum_ID)) then
		chkForumModerator = "0"
		rsUsrName.close
		exit function
	else
		MEMBER_ID = rsUsrName("MEMBER_ID")
		rsUsrName.close
	end if

	set rsUsrName = nothing

	'## Forum_SQL
	strSql = "SELECT * "
	strSql = strSql & " FROM " & strTablePrefix & "MODERATOR "
	strSql = strSql & " WHERE (FORUM_ID = " & fForum_ID & "  "
	strSql = strSql & " AND   MEMBER_ID = " & MEMBER_ID & ") OR ("
	strSQL= strSql & " MEMBER_ID = " & MEMBER_ID & " AND FORUM_ID = -1)"

	set rsChk = my_Conn.Execute (strSql)

	if rsChk.bof or rsChk.eof then
		chkForumModerator = "0"
	else
		chkForumModerator = "1"
	end if 

	rsChk.close
	set rsChk = nothing

end function

function listForumModerators(fForum_ID)

	'## Forum_SQL
	strSql = "SELECT * "
	strSql = strSql & " FROM " & strTablePrefix & "MODERATOR "
	strSql = strSql & " WHERE FORUM_ID = " & fForum_ID & " OR FORUM_ID = -1"



	set rsChk = my_Conn.Execute (strSql)

	if rsChk.EOF or not(ChkQuoteOk(fForum_ID)) then
		listForumModerators = ""
		exit function
	end if
	fMods = getMemberName(rsChk("MEMBER_ID"))
	rsChk.MoveNext
	do until rsChk.EOF
		fMods = fMods & ", " & getMemberName(rsChk("MEMBER_ID"))
		rsChk.MoveNext
	loop

	rsChk.close
	set rsChk = nothing

	listForumModerators = fMods
end function

function getMemberName(fUser_Number)

	'## Forum_SQL
	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.M_NAME"
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE MEMBER_ID = " & fUser_Number

	set rsGetMemberID = my_Conn.Execute(strSql)

	if rsGetMemberID.EOF or rsGetMemberID.BOF then
		getMemberName = ""
	else
		getMemberName = rsGetMemberID("M_NAME")
	end if

end function

function getMemberNumber(fUser_Name)

	'## Forum_SQL
	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE M_NAME = '" & fUser_Name & "'"

	set rsGetMemberID = my_Conn.Execute(strSql)

	if rsGetMemberID.EOF or rsGetMemberID.BOF then
		getMemberNumber = -1
		exit function
	end if
	getMemberNumber = rsGetMemberID("MEMBER_ID")
end function

'##############################################
'##            NT Authentication             ##
'##############################################

sub NTUser()
	if Session(strCookieURL & "username")="" then

		'## Forum_SQL 
		strSql ="SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_LEVEL, " & strMemberTablePrefix & "MEMBERS.M_PASSWORD, " & strMemberTablePrefix & "MEMBERS.M_USERNAME, " & strMemberTablePrefix & "MEMBERS.M_NAME "
		strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_USERNAME = '" & Session(strCookieURL & "userid") & "'"
		strSql = strSql & " AND   " & strMemberTablePrefix & "MEMBERS.M_STATUS = " & 1

		set rs_chk = my_conn.Execute (strSql)

		if rs_chk.BOF or rs_chk.EOF then
			strLoginStatus = 0
		else
			Session(strCookieURL & "username") = rs_chk("M_NAME")
			if strSetCookieToForum = 1 then
				Response.Cookies(strUniqueID & "User").Path = strCookieURL
			end if
			Response.Cookies(strUniqueID & "User")("Name") = rs_chk("M_NAME")
			Response.Cookies(strUniqueID & "User")("Pword") = rs_chk("M_PASSWORD")
			Response.Cookies(strUniqueID & "User")("Cookies") = ""
			Response.Cookies(strUniqueID & "User").Expires = dateAdd("d", 30, strForumTimeAdjust)	
			Session(strCookieURL & "last_here_date") = ReadLastHereDate(Request.Form("Name"))	
			if strAuthType = "nt" then		
			Session(strCookieURL & "last_here_date") = ReadLastHereDate(Session(strCookieURL & "userID"))	
			end if

			strLoginStatus = 1
			
			mLev = cint(ChkUser2(Request.Cookies(strUniqueID & "User")("Name"), Request.Cookies(strUniqueID & "User")("Pword")))
			if mLev = 4 then 
				Session(strCookieURL & "Approval") = "15916941253"
			end if
		end if

		rs_chk.close	
		set rs_chk = nothing

	end if
end sub

function ChkAccountReg()

	'## Forum_SQL 
	strSql ="SELECT " & strMemberTablePrefix & "MEMBERS.M_LEVEL, " & strMemberTablePrefix & "MEMBERS.M_USERNAME "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_USERNAME = '" & Session(strCookieURL & "userid") & "'" 
	strSql = strSql & " AND   " & strMemberTablePrefix & "MEMBERS.M_STATUS = " & 1

	set rs_chk = my_conn.Execute (strSql)

	if rs_chk.BOF or rs_chk.EOF then
		ChkAccountReg = "0"
	else
		ChkAccountReg = "1"
	end if

	rs_chk.close	
	set rs_chk = nothing

end function

sub NTAuthenticate()
	dim strUser, strNTUser, checkNT
	strNTUser = Request.ServerVariables("AUTH_USER") 
	strNTUser = replace(strNTUser, "\", "/")
	if Session(strCookieURL & "userid") = "" then
		strUser = Mid(strNTUser,(instr(1,strNTUser,"/")+1),len(strNTUser))
		Session(strCookieURL & "userid") = strUser
	end if
	if strNTGroups="1" then
		strNTGroupsSTR = Session(strCookieURL & "strNTGroupsSTR")
		if Session(strCookieURL & "strNTGroupsSTR") = "" then
			Set strNTUserInfo = GetObject("WinNT://"+strNTUser)
			For Each strNTUserInfoGroup in strNTUserInfo.Groups
				strNTGroupsSTR=strNTGroupsSTR+", "+strNTUserInfoGroup.name
			NEXT
			Session(strCookieURL & "strNTGroupsSTR") = strNTGroupsSTR
		end if
	end if

	if strAutoLogon="1" then
		strNTUserFullName = Session(strCookieURL & "strNTUserFullName")
		if Session(strCookieURL & "strNTUserFullName") = "" then
			Set strNTUserInfo = GetObject("WinNT://"+strNTUser)
			strNTUserFullName=strNTUserInfo.FullName
			Session(strCookieURL & "strNTUserFullName") = strNTUserFullName
		end if
	end if
end sub


function chkDisplayForum(fForum_ID)
	if (mlev = 4) then
		chkDisplayForum= true
		exit function
	end if

	'## Forum_SQL - load the user list       
	strSql = "SELECT " & strTablePrefix & "FORUM.F_PRIVATEFORUMS,  " & strTablePrefix & "FORUM.F_PASSWORD_NEW,  " & strTablePrefix & "FORUM.F_HIDDEN  "
	strSql = strSql & " FROM " & strTablePrefix & "FORUM "
	strSql = strSql & " WHERE FORUM_ID = " & fForum_ID

	set rsAccess = my_Conn.Execute(strSql)
	select case  rsAccess("F_PRIVATEFORUMS")

	case 0, 1, 2, 3, 4, 7, 9,10,11,12
'#########
		if rsAccess("F_HIDDEN") = 1 then
			chkDisplayForum = false
		else
		chkDisplayForum= true
		end if
		exit function
'#########
	case 5
		UserNum = getNewMemberNumber()
		if UserNum = - 1 then
			chkDisplayForum= false
			exit function
		else
			chkDisplayForum= true
			exit function
		end if
	case 6
		UserNum = getNewMemberNumber()
		if UserNum = - 1 then
			chkDisplayForum= false
			exit function
		end if
		MatchFound = isAllowedMember(fForum_ID,UserNum)
		if MatchFound = 1 then
			chkDisplayForum= true
		Else
			chkDisplayForum= false
		end if 
 	case 8
		chkDisplayForum= false
			if strAuthType="nt" THEN
				NTGroupSTR = Split(strNTGroupsSTR, ", ")
				for j = 0 to ubound(NTGroupSTR)
					NTGroupDBSTR = Split(rsAccess("F_PASSWORD_NEW"), ", ")
					for i = 0 to ubound(NTGroupDBSTR)
						if NTGroupDBSTR(i) = NTGroupSTR(j) then
							chkDisplayForum= true
							exit function
						end if
					next
				next
			End if

	case else 
			chkDisplayForum= true
	end select 

end function

'##############################################
'##        Cookie functions and Subs         ##
'##############################################

sub DoCookies(fSavePassWord)
	if strSetCookieToForum = 1 then
		Response.Cookies(strUniqueID & "User").Path = strCookieURL
	else
		Response.Cookies(strUniqueID & "User").Path = "/"
	end if
	Response.Cookies(strUniqueID & "User")("Name") = strDBNTFUserName
	Response.Cookies(strUniqueID & "User")("Pword") = Request.Form("Password")
	Response.Cookies(strUniqueID & "User")("Cookies") = Request.Form("Cookies")
	if fSavePassWord = "true" then
		Response.Cookies(strUniqueID & "User").Expires = dateAdd("d", 30, strForumTimeAdjust)
	end if
	Session(strCookieURL & "last_here_date") = ReadLastHereDate(strDBNTFUserName)	

end sub

sub ClearCookies()
	if strSetCookieToForum = 1 then
		Response.Cookies(strUniqueID & "User").Path = strCookieURL
	else
		Response.Cookies(strUniqueID & "User").Path = "/"
	end if
	Response.Cookies(strUniqueID & "User") = ""
	'Response.Cookies(strUniqueID & "User").Expires = dateadd("d", -2, strForumTimeAdjust)
end sub

'##############################################
'##                Do Counts                 ##
'##############################################

sub DoPCount()

	'## Forum_SQL - Updates the totals Table
	strSql ="UPDATE " & strTablePrefix & "TOTALS SET " & strTablePrefix & "TOTALS.P_COUNT = " & strTablePrefix & "TOTALS.P_COUNT + 1"

	my_Conn.Execute (strSql)

end sub

sub DoTCount()

	'## Forum_SQL - Updates the totals Table
	strSql ="UPDATE " & strTablePrefix & "TOTALS SET " & strTablePrefix & "TOTALS.T_COUNT = " & strTablePrefix & "TOTALS.T_COUNT + 1"

    my_Conn.Execute (strSql)

end sub

sub DoUCount(sUser_Name)

	'## Forum_SQL - Update Total Post for user
	strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " SET " & strMemberTablePrefix & "MEMBERS.M_POSTS = " & strMemberTablePrefix & "MEMBERS.M_POSTS + 1 "
	strSql = strSql & " WHERE " & strDBNTSQLName & " = '" & sUser_Name & "'"
'

	my_Conn.Execute (strSql)

end sub

'##############################################
'##              Private Forums              ##
'##############################################

sub chkUser4()

	if mLev = 4 then 
		exit sub
	end if


	'## Forum_SQL
	strSql = "SELECT " & strTablePrefix & "FORUM.F_PRIVATEFORUMS, " & strTablePrefix & "FORUM.F_SUBJECT, " & strTablePrefix & "FORUM.F_PASSWORD_NEW "
	strSql = strSql & " FROM " & strTablePrefix & "FORUM "
	strSql = strSql & " WHERE " & strTablePrefix & "FORUM.Forum_ID = " & Request.QueryString("FORUM_ID")

	set rsStatus = my_conn.Execute (strSql)

	dim Users
	If cint(rsStatus("F_PRIVATEFORUMS")) <> 0 then

			Select case cint(rsStatus("F_PRIVATEFORUMS"))
				case 0
					'## Do Nothing
				case 1, 6 '## Allowed Users
					UserNum = getNewMemberNumber()
					MatchFound = isAllowedMember(Request.QueryString("FORUM_ID"), cint(UserNum))
					if MatchFound then
						exit sub
					else
						doNotAllowed
						Response.end
					end if
				case 2 '## password
					select case Request.Cookies(strUniqueID & "User")("PRIVATE_" & rsStatus("F_SUBJECT"))
						case rsStatus("F_PASSWORD_NEW")
							'## OK
						case else
							if Request("pass") = "" then
								doPasswordForm
								Response.End
							else
								if Request("pass") <> rsStatus("F_PASSWORD_NEW") then
									Response.Write "Invalid password! <a href='" & Request.ServerVariables("HTTP_REFERER") & "'>Back</a>"
									Response.End
								else
									if strSetCookieToForum = 1 then
										Response.Cookies(strUniqueID & "User").Path = strCookieURL
									end if
									Response.Cookies(strUniqueID & "User")("PRIVATE_" & rsStatus("F_SUBJECT")) = Request("pass")
								end if
							end if
					end select
				case 3    '## Either Password or Allowed
					UserNum = getNewMemberNumber()
					MatchFound = isAllowedMember(Request.QueryString("FORUM_ID"), cint(UserNum))
					if MatchFound then
						exit sub
					else
					select case Request.Cookies(strUniqueID & "User")("PRIVATE_" & rsStatus("F_SUBJECT"))
						case rsStatus("F_PASSWORD_NEW")
							'## OK
						case else
							if Request("pass") = "" then
								doLoginForm
								Response.End
							else
								if Request("pass") <> rsStatus("F_PASSWORD_NEW") then
									Response.Write "Invalid password! <a href='" & Request.ServerVariables("HTTP_REFERER") & "'>Back</a>"
									Response.End
								else
									if strSetCookieToForum = 1 then
										Response.Cookies(strUniqueID & "User").Path = strCookieURL
									end if
									Response.Cookies(strUniqueID & "User")("PRIVATE_" & rsStatus("F_SUBJECT")) = Request("pass")
								end if
							end if
					end select
					end if
				'## code added 07/13/2000
				case 7    '## members or password
					if (strDBNTUserName = "") then
						select case Request.Cookies(strUniqueID & "User")("PRIVATE_" & rsStatus("F_SUBJECT"))
							case rsStatus("F_PASSWORD_NEW")
								'## OK
							case else
								if Request("pass") = "" then
									doLoginForm
									Response.End
								else
									if Request("pass") <> rsStatus("F_PASSWORD_NEW") then
										Response.Write "Invalid password! <a href='" & Request.ServerVariables("HTTP_REFERER") & "'>Back</a>"
										Response.End
									else
										if strSetCookieToForum = 1 then
											Response.Cookies(strUniqueID & "User").Path = strCookieURL
										end if
										Response.Cookies(strUniqueID & "User")("PRIVATE_" & rsStatus("F_SUBJECT")) = Request("pass")
									end if
								end if
						end select
					end if
				'## end code added 07/13/2000
				case 4, 5 '## members only
					if strDBNTUserName = "" then
						doNotLoggedInForm
					end if
				case 8, 9
					NTGroupSTR = Split(strNTGroupsSTR, ", ")
					NTGroupDBSTR = Split(rsStatus("F_PASSWORD_NEW"), ", ")
						For i = 0 to ubound(NTGroupDBSTR)
							for j = 0 to ubound(NTGroupSTR)
								if NTGroupDBSTR(i) = NTGroupSTR(j) then
									exit SUB
								end if
							next
						next
					doNotAllowed
					Response.end
'############### READ/WRITE ACCESS ###########################
				case 10,11,12
'############### READ/WRITE ACCESS ###########################				
				case else
					Response.Write "<BR>ERROR: Invalid forum type: " & rsStatus("F_PRIVATEFORUMS")
					Response.end
			end select
	end if

	'my_Conn.Close
	'set my_Conn = nothing

end sub

function chkForumAccess(fForum)

	if mLev = 4 then 
		chkForumAccess = true
		exit function
	end if


	'## Forum_SQL
	strSql = "SELECT " & strTablePrefix & "FORUM.F_PRIVATEFORUMS, " & strTablePrefix & "FORUM.F_SUBJECT, " & strTablePrefix & "FORUM.F_PASSWORD_NEW "
	strSql = strSql & " FROM " & strTablePrefix & "FORUM "
	strSql = strSql & " WHERE " & strTablePrefix & "FORUM.Forum_ID = " & fForum

	set rsStatus = my_conn.Execute (strSql)

	dim Users
	dim MatchFound

	If cint(rsStatus("F_PRIVATEFORUMS")) <> 0 then

			Select case cint(rsStatus("F_PRIVATEFORUMS"))
				case 0
					chkForumAccess = true
				case 1, 6 '## Allowed Users
					UserNum = getNewMemberNumber()
					chkForumAccess = (isAllowedMember(fForum,UserNum) = 1)
				case 2 '## password
					select case Request.Cookies(strUniqueID & "User")("PRIVATE_" & rsStatus("F_SUBJECT"))
						case rsStatus("F_PASSWORD_NEW")
							chkForumAccess = true
						case else
							if Request("pass") = "" then
								chkForumAccess = false
							else
								if Request("pass") <> rsStatus("F_PASSWORD_NEW") then
									chkForumAccess = false
								else
									if strSetCookieToForum = 1 then
										Response.Cookies(strUniqueID & "User").Path = strCookieURL
									end if
									Response.Cookies(strUniqueID & "User")("PRIVATE_" & rsStatus("F_SUBJECT")) = Request("pass")
									chkForumAccess = true
								end if
							end if
					end select
				case 3    '## Either Password or Allowed
					UserNum = getNewMemberNumber()
				'	if countMembers(fForum) = 0 then
				'		chkForumAccess = false 
				'		exit function
				'	end if					
					if isAllowedMember(fForum,UserNum) = 1 then
						chkForumAccess = true
					else
						chkForumAccess = false
					end if
					if not(chkForumAccess) then 
						select case Request.Cookies(strUniqueID & "User")("PRIVATE_" & rsStatus("F_SUBJECT"))
							case rsStatus("F_PASSWORD_NEW")
								chkForumAccess = true
							case else
								if Request("pass") = "" then
									chkForumAccess = false
								else
									if Request("pass") <> rsStatus("F_PASSWORD_NEW") then
										chkForumAccess = false
									else
										if strSetCookieToForum = 1 then
											Response.Cookies(strUniqueID & "User").Path = strCookieURL
										end if
										Response.Cookies(strUniqueID & "User")("PRIVATE_" & rsStatus("F_SUBJECT")) = Request("pass")
										chkForumAccess = true
									end if
								end if
						end select
					end if
				'## code added 07/13/2000
				case 7    '## members or password
					if strDBNTUserName = "" then
						select case Request.Cookies(strUniqueID & "User")("PRIVATE_" & rsStatus("F_SUBJECT"))
							case rsStatus("F_PASSWORD_NEW")
								chkForumAccess = true
							case else
								if Request("pass") = "" then
									chkForumAccess = false
								else
									if Request("pass") <> rsStatus("F_PASSWORD_NEW") then
										chkForumAccess = false
									else
										if strSetCookieToForum = 1 then
											Response.Cookies(strUniqueID & "User").Path = strCookieURL
										end if
										Response.Cookies(strUniqueID & "User")("PRIVATE_" & rsStatus("F_SUBJECT")) = Request("pass")
										chkForumAccess = true
									end if
								end if
						end select
					end if
				'## end code added 07/13/2000						
					
				case 4, 5 '## members only
					if strDBNTUserName = "" then
						chkForumAccess = false
					else 				'V3.1 SR4
						chkForumAccess = true
					end if

				case 8, 9   
					test="test db"
					chkForumAccess = FALSE
					if strAuthType="db" then
						chkForumAccess = true
						exit function
					end if              
				NTGroupSTR = Split(strNTGroupsSTR, ", ")
				for j = 0 to ubound(NTGroupSTR)
					NTGroupDBSTR = Split(rsStatus("F_PASSWORD_NEW"), ", ")
					for i = 0 to ubound(NTGroupDBSTR)
						if NTGroupDBSTR(i) = NTGroupSTR(j) then
							chkForumAccess = True    
							exit function
						end if
					next
				next       

				case else    
					chkForumAccess = true
			end select
	else
		chkForumAccess = true
	end if

end function


function chkAccess(fForum)

	if mLev = 4 then 
		chkAccess = true
		exit function
	end if
	
	'## Forum_SQL - load the user list
	strSql = "SELECT " & strTablePrefix & "FORUM.F_PRIVATEFORUMS FROM " & strTablePrefix & "FORUM WHERE FORUM_ID = " & fForum

	set rsAccess = my_Conn.Execute(strSql)

	if rsAccess("F_PRIVATEFORUMS") <> 1 then
		chkAccess = true
		exit function
	end if
	if Request.Cookies(strUniqueID & "User")("Name") = "" then
		chkAccess = false
	end if
	'get the member number
	UserNum = getMemberNumber(Request.Cookies(strUniqueID & "User")("Name"))
	if isAllowedMember(fForum,UserNum) = 1 then
		chkAccess = true
	else
		chkAccess = false
	end if
End function

sub doLoginForm()
%>
<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">There Was A Problem</font></p>

<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">
You do not have access to this forum.
</font></p>

<p align=center><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>>If you have been given special permission by the administrator to view and/or post in this forum, enter the password here:
<form action="<% =Request.ServerVariables("SCRIPT_NAME") %>" id=form2 name=form2>
<%
	for each q in Request.QueryString
		Response.Write "<input type=hidden name=""" & q & """ value=""" & Request(q) & """>"
	next
%>
<input name=pass type=password>
<input type=submit value=Enter id=submit2 name=submit2>
</form></p>

<p align=center><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>><a href="JavaScript:history.go(-1)">请返回重新输入</a></font></p>
<p align=center><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>><a href="default.asp">Return to the forum</a></font></p>
<!--#INCLUDE FILE="inc_footer.asp"-->
<%
end sub

sub doNotAllowed()
%>
<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">There Was A Problem</font></p>

<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">
You do not have access to this forum.
</font></p>

<p align=center><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>><a href="JavaScript:history.go(-1)">请返回重新输入</a></font></p>
<p align=center><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>><a href="default.asp">Return to the forum</a></font></p>
<!--#INCLUDE FILE="inc_footer.asp"-->
<%
end sub

sub doPasswordForm()
%>
<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">There Was A Problem</font></p>

<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">
You must enter the password for this forum.

<form action="<% =Request.ServerVariables("SCRIPT_NAME") %>" id=form2 name=form2>
<%
	for each q in Request.QueryString
		Response.Write "<input TYPE=hidden name=""" & q & """ value=""" & Request(q) & """>"
	next
%>
<input name=pass type=password>
<input type=submit value=Enter id=submit1 name=submit1>
</form>
</font></p>

<p align=center><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>><a href="JavaScript:history.go(-1)">请返回重新输入</a></font></p>

<p align=center><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>><a href="default.asp">Return to the forum</a></font></p>
<!--#INCLUDE FILE="inc_footer.asp"-->
<%
end sub

sub doNotLoggedInForm()
%>
<p align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">There Was A Problem</font></p>

<P align=center><font face="<% =strDefaultFontFace %>" size="<% =strHeaderFontSize %>">
You are not logged in.
</font></p>

<p align=center><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>><a href="JavaScript:history.go(-1)">请返回重新输入</a></font></p>
<p align=center><font face="<% =strDefaultFontFace %>" size=<% =strDefaultFontSize %>><a href="default.asp">Return to the forum</a></font></p>
<!--#INCLUDE FILE="inc_footer.asp"-->
<%
	Response.End
end sub

function getNewMemberNumber()

dim my_Conn2

	set my_Conn2 = Server.CreateObject("ADODB.Connection")
	my_Conn2.Open strConnString

	'## Forum_SQL
	strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE " & strDBNTSQLName & " = '" & strDBNTUserName & "'"
 	set rsGetMemberID = my_Conn2.Execute(strSql)
	if rsGetMemberID.EOF or rsGetMemberID.BOF then
		getNewMemberNumber = -1
		exit function
	end if
	getNewMemberNumber = rsGetMemberID("MEMBER_ID")
	my_Conn2.Close
	Set my_Conn2 = nothing
end function


Function ReplaceUrls(fString)
	Dim oTag, c1Tag, c2Tag
	Dim roTag, rc1Tag, rc2Tag
	Dim oTagPos, c1TagPos, c2TagPos
	Dim nTagPos
	Dim counter2
	Dim strArray, strArray2, strArray3

	oTag   = "[url="""
    oTag2  = "[url]"
    roTag  = "<a href="""
    c1Tag  = """]"
    c1Tag2 = "[/url]"
    rc1Tag = """ target=""_New"">"
    c2Tag  = "[/url]"
    rc2Tag = "</a>"
    oTagPos = InStr(1, fString, oTag, 1)
    c1TagPos = InStr(1, fString, c1Tag, 1)
   
strTempString = ""
if (oTagpos > 0) and (c1TagPos > 0) then
	strArray = Split(fString, oTag, -1)

	for counter2 = 0 to UBound(strArray)
		if (InStr(1, strArray(counter2), c2Tag, 1) > 0) and (InStr(1, strArray(counter2), c1Tag, 1) > 0) then
			strArray2 = Split(strArray(counter2), c1Tag, -1)
			if Instr(1, strArray2(1), c2Tag) and  not( (Instr(1, UCase(strArray2(1)), "[URL]") >0) and not(Instr(1, UCase(strArray2(1)), "[/URL]") >0) ) then
'			if Instr(1, strArray2(1), c2Tag) then  
				strFirstPart = Left(strArray2(1), Instr(1, strArray2(1),c2Tag)-1)
				strSecondPart = Right(strArray2(1), (Len(strArray2(1)) - Instr(1, strArray2(1), c2Tag) - len(c2Tag)+1))
				if strFirstPart <> "" then
					if (Instr(strArray2(0),"@") > 0) and UCase(Left(strArray2(0), 7)) <> "MAILTO:" then
						strTempString = strTempString & roTag & "mailto:" & strArray2(0) & rc1Tag & strFirstPart & rc2Tag & strSecondPart
					else
						strTempString = strTempString & roTag & strArray2(0) & rc1Tag & strFirstPart & rc2Tag & strSecondPart
					end if
				else
					if (Instr(strArray2(0),"@") > 0) and UCase(Left(strArray2(0), 7)) <> "MAILTO:" then
						strTempString = strTempString & roTag & "mailto:" & strArray2(0) & rc1Tag & strArray2(0) & rc2Tag & strSecondPart
				else
					strTempString = strTempString & roTag & strArray2(0) & rc1Tag & strArray2(0) & rc2Tag & strSecondPart
				end if
				end if
			else
				strTempString = strTempString & roTag & strArray2(0) & rc1Tag & strArray2(0) & rc2Tag & strArray2(1)
			end if
		elseif (InStr(1, strArray(counter2), c1Tag, 1) > 0) then
			strArray2 = Split(strArray(counter2), c1Tag, -1)
			strTempString = strTempString & roTag & strArray2(0) & rc1Tag & strArray2(0) & rc2Tag & strArray2(1)
		else
			strTempString = strTempString & strArray(counter2)
		end if
	next

else
	strTempString = fString
end if

oTagPos2 = InStr(1, strTempString, oTag2, 1)
c1TagPos2 = InStr(1, strTempString, c1Tag2, 1)

if (oTagpos2 > 0) and (c1TagPos2 > 0) then
 	strTempString2 = ""
 	strArray = Split(strTempString, oTag2, -1)
 	for counter3 = 0 to Ubound(strArray)
 		if (Instr(1, strArray(counter3), c1Tag2) > 0) then
 			strArray2 = split(strArray(counter3), c1Tag2, -1)
			if (Instr(strArray2(0),"@") > 0) and UCase(Left(strArray2(0), 7)) <> "MAILTO:" then
	 			strTempString2 = strTempString2 & roTag & "mailto:" & strArray2(0) & rc1Tag & strArray2(0) & rc2Tag & strArray2(1)
			else
	 			strTempString2 = strTempString2 & roTag & strArray2(0) & rc1Tag & strArray2(0) & rc2Tag & strArray2(1)
			end if
 		else
 			strTempString2 = strTempString2 & strArray(counter3)
 		end if	
 	next  
 	strTempString = strTempString2
end if

	ReplaceUrls = strTempString

end function

function isAllowedMember(fForum_ID,fMemberID)

		isAllowedMember = 0
		on error resume next
		strSql = "SELECT MEMBER_ID, FORUM_ID FROM " & strTablePrefix & "ALLOWED_MEMBERS "
		strSql = strSql & " WHERE " & strTablePrefix & "ALLOWED_MEMBERS.FORUM_ID = " & fForum_ID
		strSql = strSql & " AND " & strTablePrefix & "ALLOWED_MEMBERS.MEMBER_ID = " & fMemberID

		set rsAllowedMember = my_Conn.execute (strSql)
		if (rsAllowedMember.EOF or rsAllowedMember.BOF) then
			isAllowedMember = 0
			exit function
		else
			isAllowedMember = 1
		end if

end function
%>
<script language="javascript1.2" runat=server>
function edit_hrefs(s_html, type){
    s_str = new String(s_html);
	if (type == 1) {
     	s_str = s_str.replace(/\b(http\:\/\/[\w+\.]+[\w+\.\:\/\_\?\=\&\-\'\#\%\~\;\,\$\!\+\*]+)/gi,
		  "<a href=\"$1\" target=\"_blank\">$1<\/a>");
	} 
	if (type == 2) {

		s_str = s_str.replace(/\b(https\:\/\/[\w+\.]+[\w+\.\:\/\_\?\=\&\-\'\#\%\~\;\,\$\!\+\*]+)/gi,
		  "<a href=\"$1\" target=\"_blank\">$1<\/a>");
	}
	if (type == 3) {
		s_str = s_str.replace(/\b(file\:\/\/\/\w\:\\[\w+\/\w+\.\:\/\_\?\=\&\-\'\#\%\~\;\,\$\!\+\*]+)/gi,
		  "<a href=\"$1\" target=\"_blank\">$1<\/a>");
	}
	if (type == 4) {

		s_str = s_str.replace(/\b(www\.[\w+\.\:\/\_\?\=\&\-\'\#\%\~\;\,\$\!\+\*]+)/gi,
 		  "<a href=\"http://$1\" target=\"_blank\">$1</a>");
	}
	if (type == 5) {
		s_str = s_str.replace(/\b([\w+\-\'\#\%\.\_\,\$\!\+\*]+@[\w+\.?\-\'\#\%\~\_\.\;\,\$\!\+\*]*)/gi,
 		  "<a href=\"mailto\:$1\">$1</a>");
	}
		  	  
    return s_str;
}
</script> 
<% 
'##############################################
'##                 My Suff                  ##
'##############################################

'Rem User Field Code #######################################
function DoesUserFieldExist(fMemberID,FUSerFieldID)

	fieldSQL = "SELECT * FROM " & strMemberTablePrefix & "MEMBERFIELDS WHERE " & strMemberTablePrefix & "MEMBERFIELDS.USR_FIELD_ID=" & FUSerFieldID 
	fieldSQL = fieldSQL & " AND " & strMemberTablePrefix & "MEMBERFIELDS.MEMBER_ID=" & fMemberID
	set rsGetFields = my_Conn.Execute(fieldSQL)
	DoesUserFieldExist = not rsGetFields.EOF
	rsGetFields.Close

end function

function getUserFieldValue(fUser_Field_ID,fMemberID)

	strSql = "SELECT " & strMemberTablePrefix & "MEMBERFIELDS.USR_VALUE "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERFIELDS "
	strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERFIELDS.USR_FIELD_ID = " & fUser_Field_ID
	strSql = strSql & " AND " & strMemberTablePrefix & "MEMBERFIELDS.MEMBER_ID = " & fMemberID 
	set rsgetUserFieldValue = my_Conn.Execute(strSql)
	if Not rsgetUserFieldValue.EOF then
		getUserFieldValue = rsgetUserFieldValue("USR_VALUE")
	else
		getUserFieldValue = ""
	end if

end function
'Rem User Field Code #######################################


function SpellcheckASP(TXTTOSC)
on error resume next
set sc = GetObject("java:SpellCheck")
if err.number <> 0 then
response.write "Java Class Not installed"
response.end
end if
sc.loadDictionary Server.MapPath("tools/dictionary.txt")
sc.setHighlights "<u><b>", "</b></u>"
SpellcheckASP=sc.checkSpelling(TXTTOSC)
end function

Function sGetColspan(lIN, lOUT)	
if (strShowModerators = "1") Then lOut = lOut + 1	
If (mlev = "4" or mlev = "3") then lOut = lOut + 1		
If lOut > lIn then
	sGetColspan = lIN
Else
	sGetColspan = lOUT	
End If
End Function

'##############################################
function remoteIP()	
	
	remoteIP = Request.ServerVariables("HTTP_X_FORWARDED_FOR")	
	if remoteIP = "" then		
		remoteIP = Request.ServerVariables("REMOTE_ADDR")	
	end if
end function
 %>
 
