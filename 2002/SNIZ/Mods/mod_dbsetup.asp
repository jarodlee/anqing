<!--#include file="../config.asp"-->
<!--#include file="../inc_functions.asp"-->

<%     
	Dim strTableName
	Dim fieldArray (50)
	Dim idFieldName
	Dim ErrorCount
	
	ErrorCount = 0
	
	set my_Conn= Server.CreateObject("ADODB.Connection")
	my_Conn.Open strConnString

	Set fso = CreateObject("Scripting.FileSystemObject")
    set objFile = fso.Getfile(server.mappath(modFolder & "mod_dbsetup.asp"))
    set objFolder = objFile.ParentFolder
    set objFolderContents = objFolder.Files

if Request.Form("dbMod") = "" then
	Response.write("<font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>")
	Response.write "<h3>Snitz Forum Modifications</h3>" &_
	"<b>" & strDBType & "</b>&nbsp;Database Setup.....<br></font>" &_
	"<form action=""mod_dbsetup.asp"" method=""post"" name=""form1"">"
 	
	If strDBType = "" then 
		Response.write "<font face=""" & strDefaultFontFace & """ color=red size=""" & strDefaultFontSize & """>Your strDBType is not set, please edit your config.asp<br>" &_
		"to reflect your database type<br></font>"
		Response.write "<br><a href=""../default.asp"">返回讨论区</a></font>"
		Response.End
	end if
	If strDBType = "sqlserver" then 
		Response.write "<font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """>" &_
		"You are using SQL Server, please select the correct version<br>" &_
		"<input type=""radio"" name=""sqltype"" value=""7"" checked> SQL 7.x&nbsp;&nbsp;&nbsp;&nbsp;" &_
		"<input type=""radio"" name=""sqltype"" value=""6""> SQL 6.x<br></font>"
 	End If 
	on error resume next
	Response.write "<font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """>" &_
	"<p>Select the Mod from the list below, and press Update!<br>" &_
	"A script will execute to perform the database upgrade.</p></font>" &_
	"<select name=""dbMod"" size=""1"">"
	for each objFileItem in objFolderContents
		intFile = instr(objFileItem.Name, ".dbs")
	    if intFile <> 0 then
	        modFile = left(objFileItem.Name, intFile - 1)
	        whichfile=server.mappath(modFolder & modFile & ".dbs")

		    Set fs = CreateObject("Scripting.FileSystemObject")
	        Set thisfile = fs.OpenTextFile(whichfile, 1, False)
			ModName = thisfile.readline
			Response.write "<option value=""" & whichfile & """>" & ModName & "</option>"
			thisfile.close
			if err.number <> 0 then 
				Response.write err.description
				Responce.end
			end if
			set fs = nothing
	  	end if
	Next
	Response.write "</select>" &_
	"<input type=""submit"" name=""submit1"" value=""Update!""></form>" &_
	"<font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" &_
	"<a href=""../default.asp"">返回讨论区</a></font>"
	
 Else
	sqlVer = Request.Form("sqltype")
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set thisfile = fs.OpenTextFile(Request.Form("dbMod"), 1, False)
	ModName = thisfile.readline
	response.write ("<font face=" & strDefaultFontFace & " size=" & strDefaultFontSize+1 & ">")
	response.write ("<h4>" & ModName & "</h4></font>")
	sectionName     =   thisfile.readline
    sectionNo = 0
	maxrec = 0

'Load Create Tables
	do while not thisfile.AtEndOfStream
	sectionName = thisfile.readline
	Select case sectionName
		case "[CREATE]" 
			strTableName   =   thisfile.readline
			idFieldName = thisfile.readline
			tempField = thisfile.readline
			rec = 0
			do while tempField <>  "[END]"
				fieldArray(rec) = tempField
				rec = rec+1
				tempField = thisfile.readline
			loop
			CreateTables(rec)
			sectionNo = 0
			maxrec = 0
		case "[ALTER]" 
			strTableName   =   thisfile.readline
			tempField = thisfile.readline
			rec = 0
			do while tempField <>  "[END]"
				fieldArray(rec) = tempField
				rec = rec+1
				tempField = thisfile.readline
			loop
			AlterTables(rec)
		case "[DELETE]" 
			strTableName   =   thisfile.readline
			tempField = thisfile.readline
			rec = 0
			do while tempField <>  "[END]"
				fieldArray(rec) = tempField
				rec = rec+1
				tempField = thisfile.readline
			loop
			DeleteRecs(rec)
		case "[INSERT]" 
			strTableName   =   thisfile.readline
			tempField = thisfile.readline
			rec = 0
			do while tempField <>  "[END]"
				fieldArray(rec) = tempField
				rec = rec+1
				tempField = thisfile.readline
			loop
			InsertValues(rec)
		case "[UPDATE]" 
			strTableName   =   thisfile.readline
			tempField = thisfile.readline
			rec = 0
			do while tempField <>  "[END]"
				fieldArray(rec) = tempField
				rec = rec+1
				tempField = thisfile.readline
			loop
			UpdateValues(rec)
	end select
	
	loop
	response.write ""

	Response.write "<br><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><p><b>Database setup finished</b></p>"
	if ErrorCount > 0 then
		Response.write "There were errors.....<br>"
	end if
	Response.write "</font>" &_
	"<form action=""mod_dbsetup.asp"" method=""post"" name=""form2"">" &_
	"<input type=""hidden"" name=""dbMod"" value="""">" &_
	"<input type=""submit"" name=""submit2"" value=""Finished""></form>"
end if 

set fs=nothing
set fso = nothing

Sub CreateTables( numfields )
Response.write "<br><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """>"
response.write "<b>Creating table(s)...</b><br>"

	strSql = "CREATE TABLE " & strTablePrefix & strTableName & "( "
	if idFieldName <> "" then
		select case strDBType
			case "access"
				if Instr(strConnString,"(*.mdb)") then
					strSql = strSql & idFieldName &" int COUNTER NOT NULL "
				else
					strSql = strSql & idFieldName &" int IDENTITY (1, 1) NOT NULL "
				end if
			case "sqlserver"
				strSql = strSql & idFieldName &" int IDENTITY (1, 1) NOT NULL "
			case "mysql"
				strSql = strSql & idFieldName &" INT (11) DEFAULT '' NOT NULL auto_increment "
		end select
	end if
	for y = 0 to numfields -1
	on error resume next
		tmpArray = split(fieldArray(y),";")
		fName = tmpArray(0)
		fType = tmpArray(1)
		fNull = tmpArray(2)
		fDefault = tmpArray(3)
		if idFieldName <> "" or y <> 0 then
			strSql = strSql & ", "
		end if
		select case strDBType
			case "access"
				
				fType = replace(fType,"varchar (","text (")
			case "sqlserver"
				select case sqlVer
					case 7
						fType = replace(fType,"memo","ntext")
						fType = replace(fType,"varchar","nvarchar")
					case else
						fType = replace(fType,"memo","text")
				end select
			case "mysql"
				fType = replace(fType,"memo","text")
				fType = replace(fType,";int",";int (11)")
				fType = replace(fType,";smallint",";smallint (6)")
		end select
		if fNull <> "NULL" then fNull = "NOT NULL"
		strSql = strSql & fName & " " & fType & " " & fNull & " " 
		if fdefault <> "" then
			strSql = strSql & "DEFAULT " & fDefault
		end if
	next
	if strDBType = "mysql" then
		strSql = strSql & ",KEY " & strTablePrefix & strTableName & "_" & idFieldName & "(" & idFieldName & "))"
	else
		strSql = strSql & ")"
	end if
	response.write strSql & "<br>"
	my_Conn.Execute strSql
	if err.number <> 0 then
		response.write("<font color=red>" & err.number & " | " & err.description & "</font><br>")
		ErrorCount = ErrorCount + 1
	else
		response.write("<b>Table created succesfully</b><br>")
	end if

	response.write("<hr size=""1"" width=""260"" align=left color=""blue""></font>")
end Sub

Sub AlterTables(numfields)
Response.write "<br><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """>"
for y = 0 to numfields -1
	on error resume next
	if strTableName = "MEMBERS" then
		TablePrefix = strMemberTablePrefix
	else
		TablePrefix = strTablePrefix
	end if
	strSql = "ALTER TABLE " & TablePrefix & strTableName & " ADD "

	if strDBType = "access" then strSql = strSql & "COLUMN "
	
		tmpArray = split(fieldArray(y),";")
		fName = tmpArray(0)
		fType = tmpArray(1)
		fNull = tmpArray(2)
		fDefault = tmpArray(3)
		select case strDBType
			case "access"
				fType = replace(fType,"text","memo")
				fType = replace(fType,"varchar","text")
			case "sqlserver"
				select case sqlVer
					case 7
						fType = replace(fType,"text","ntext")
						fType = replace(fType,"varchar","nvarchar")
				end select
			case "mysql"
				fType = replace(fType,";int",";int (11)")
				fType = replace(fType,";smallint",";smallint (6)")
		end select
		if fNull <> "NULL" then fNull = "NOT NULL"
		strSql = strSQL & fName & " " & fType & " " & fNULL & " "
		if fDefault <> "" then strSQL = strSQL & "DEFAULT " & fDefault
		
		response.write "<b>Adding Column...</b><br>" & strSQL & "<br>"
		my_Conn.Execute strSql
	if err.number <> 0 then
		response.write("<font color=red>" & err.number & " | " & err.description & "</font><br>")
		ErrorCount = ErrorCount + 1
		else
			response.write("<b>Column added succesfully</b><br>")
			if fDefault <> "" then
				strSQL = "UPDATE " & TablePrefix & strTableName & " SET " & fName & "=" & fDefault
				my_Conn.Execute strSql
				response.write "<b>Updating Current Records </b>" & strSQL & "<br>"
			end if
		end if
		if fieldArray(y) = "" then y = numfileds
	next
response.write("<b>Table(s) updated</b><br>")
response.write("<hr size=""1"" width=""260"" align=left color=""blue""></font>")
end Sub

Sub InsertValues(numfields)
Response.write "<br><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """>"
on error resume next
	response.write ("<b>Adding new records..</b><br>")
	for y = 0 to numfields-1
		if strTableName = "MEMBERS" then
		strSql = "INSERT INTO " & strMemberTablePrefix & strTableName(x) & " "
		else
		strSql = "INSERT INTO " & strTablePrefix & strTableName & " "
		end if
		tmpArray = split(fieldArray(y),";")
		fNames = tmpArray(0)
		fValues = tmpArray(1)
		strSql = strSql & tmpArray(0) & " VALUES " & tmpArray(1)
		my_Conn.Execute strSql
		response.write strSql & "<br>"
	next

	if err.number <> 0 then
		response.write("<font color=red>" & err.number & " | " & err.description & "</font><br>")
		ErrorCount = ErrorCount + 1
	else
		response.write("<br><b>Value(s) updated succesfully</b>")
	end if
response.write("<hr size=""1"" width=""260"" align=left color=""blue""><font>")
end Sub 

Sub UpdateValues(numfields)
on error resume next
Response.write "<br><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """>"
	response.write ("<b>Updating Forum Values..</b><br>")
	for y = 0 to numfields-1
		if strTableName = "MEMBERS" then
		strSql = "UPDATE " & strMemberTablePrefix & strTableName & " SET"
		else
		strSql = "UPDATE " & strTablePrefix & strTableName & " SET"
		end if
		tmpArray = split(fieldArray(y),";")
		fName = tmpArray(0)
		fValue = tmpArray(1)
		fwhere = tmpArray(2)
		strSql = strSql & " " & fName & " = " & fvalue & " WHERE " & fWhere
		my_Conn.Execute strSql
		response.write strSql & "<br>"
	next

	if err.number <> 0 then
		response.write("<font color=red>" & err.number & " | " & err.description & "</font><br>")
		ErrorCount = ErrorCount + 1
	else
		response.write("<br><b>Value(s) updated succesfully</b>")
	end if
response.write("<hr size=""1"" width=""260"" align=left color=""blue""><font>")
end Sub 

Sub DeleteValues(numfields)
on error resume next
Response.write "<br><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """>"
	response.write ("<b>Updating Forum Values..</b><br>")
		if strTableName = "MEMBERS" then
		strSql = "DELETE FROM " & strMemberTablePrefix & strTableName & " WHERE "
		else
		strSql = "DELETE FROM " & strTablePrefix & strTableName & " WHERE "
		end if
		tmpArray = fieldArray(0)
		strSql = strSql & tmpArray
		my_Conn.Execute strSql
		response.write strSql & "<br>"

	if err.number <> 0 then
		response.write("<font color=red>" & err.number & " | " & err.description & "</font><br>")
		ErrorCount = ErrorCount + 1
	else
		response.write("<br><b>Value(s) updated succesfully</b>")
	end if
response.write("<hr size=""1"" width=""260"" align=left color=""blue""><font>")
end Sub 

my_Conn.Close
set my_Conn = nothing
%>