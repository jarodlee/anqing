<br>
<table width="100%" border="0" cellspacing="1" cellpadding="4" align="center" bgcolor="<% =strTableBorderColor %>">
  <tr bgcolor="<%= strCategoryCellColor %>"> 
    <td colspan="6" height="30"> 
      <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" size"<% =strDefaultFontSize %>"><b>后加辅助功能管理</b></font>
    </td>
  </tr>
  <tr align="center" bgcolor="<%= strHeadCellColor %>"> 
    <td> 
      <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" size"<% =strDefaultFontSize %>">名称</font>
    </td>
    <td> 
      <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" size"<% =strDefaultFontSize %>">管理</font>
    </td>
    <td> 
      <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" size"<% =strDefaultFontSize %>">设定</font>
    </td>
    <td> 
      <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" size"<% =strDefaultFontSize %>">阅读</font>
    </td>
    <td> 
      <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" size"<% =strDefaultFontSize %>">版本</font>
    </td>
    <td> 
      <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>" color="<% =strCategoryFontColor %>" size"<% =strDefaultFontSize %>">作者</font>
    </td>
  </tr>

<%
    Dim fs, thisfile, thisline, whichfile, modFile, modFolder
    Dim modName, modLink, modReadme, modVersion, modSetup, modAuthor
    Dim objFile, objFolder, objFolderContents, objFileItem

    modFolder = "mods/"
    Set fso = CreateObject("Scripting.FileSystemObject")

    set objFile = fso.Getfile(server.mappath(modFolder & "modCMD.asp"))
    set objFolder = objFile.ParentFolder
    set objFolderContents = objFolder.Files

    for each objFileItem in objFolderContents
        modFile = instr(objFileItem.Name, ".mod")
        if modFile <> 0 then
            modFile = left(objFileItem.Name, modfile - 1)
            whichfile=server.mappath(modFolder & modFile & ".mod")

            Set fs = CreateObject("Scripting.FileSystemObject")
            Set thisfile = fs.OpenTextFile(whichfile, 1, False)

            modName     =   "&nbsp;" & thisfile.readline
            modLink     =   thisfile.readline
            if modLink = "" then
                modLink = "-"
            else
                modLink = "<a href="""  & modLink & """>执行</a>"
            end if
            if thisfile.readline = "false" then
                modReadme = "-"
            else
                modReadme = "<a href="""  & modFolder & "readme.asp?readme="
                modReadme = modReadme & server.urlencode(modFile) & """>查看</a>"
            end if
            modVersion  =   thisfile.readline
			if thisfile.AtEndOfStream then
				modSetup 	=	"-"
				modAuthor	=	"-"
			else
				modSetup	=	thisfile.readline
    	        if modSetup = "" then
        	        modSetup = "-"
            	else
                	modSetup = "<a href="""  & modSetup & """>go</a>"
	            end if
				modAuthor	=	thisfile.readline
			end if

            thisfile.Close
            set fs=nothing
%> 

  <tr bgcolor="<%= strForumCellColor %>" align="center"> 
    <td align="left"> 
      <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><%= modName %></font>
    </td>
    <td> 
      <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><%= modLink %></font>
    </td>
    <td> 
      <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><%= modSetup %></font>
    </td>
    <td> 
      <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><%= modReadme %></font>
    </td>
    <td> 
      <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><%= modVersion %></font>
    </td>
    <td> 
      <font face="<% =strDefaultFontFace %>" size="<% =strDefaultFontSize %>"><%= modAuthor %></font>
    </td>
  </tr>
  
<%
        end if
    Next     
    set fso = nothing
%>

</table>