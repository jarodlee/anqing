<!-- #include file="clsCollection.inc" -->
<!-- #include file="Encryption.inc" -->
<!-- #include file="clsFile.inc" -->
<!-- #include file="TypeConvert.inc" -->
<%
Class clsUpload

 Private m_objFiles	'Collection object to store given binary files
 Private m_objForm	'Collection object to store given Form Key/Value pairs

 '-------------------------------------------------------------------------------
 '- Event : Class_Initialize
 '-------------------------------------------------------------------------------
 Private Sub Class_Initialize()
   Set m_objFiles = new clsCollection
   Set m_objForm = new clsCollection
 End Sub

 '-------------------------------------------------------------------------------
 '- Property : Form (Get)
 '-------------------------------------------------------------------------------
 Public Property Get Form()
   Set Form = m_objForm
 End Property

 '-------------------------------------------------------------------------------
 '- Property : Files (Get)
 '-------------------------------------------------------------------------------
 Public Property Get Files()
   Set Files = m_objFiles
 End Property

 '-------------------------------------------------------------------------------
 '- Function    : ParseRequest
 '- Description : Parses the document parameters
 '-------------------------------------------------------------------------------
 Public Sub ParseRequest()
 Dim lngTotalBytes	'Total No. of bytes in request
 Dim lngPosBeg		'Start position for parameters in request
 Dim lngPosEnd		'End position for parameters in request
 Dim lngHelpPos		'Help variable for string positions
 Dim strBinRequest	'Binary representation of the request

 Dim strBinBoundary	'Binary boundary string when form is submitted with
                        'ENCTYPE="multipart/form-data"
 Dim lngPosBoundary	'Start position of boundary

 Dim strParameter	'Parameter Key/Value pair
 Dim strName		'Parameter Key
 Dim strValue		'Parameter Value
 Dim lngPosDelimiter	'Position of delimiter in Key/Value pair
 Dim strFileName	'FileName in Form
 Dim strContentType	'Content type in Form
 Dim strBinContent	'Form content in Binary format
 Dim objFile            'clsFile object for Binary files in Form
 Dim lngPosFileName
   lngTotalBytes = Request.TotalBytes
   strBinRequest = Request.BinaryRead(lngTotalBytes)
   lngPosBeg = 1
   lngPosEnd = InstrB(lngPosBeg, strBinRequest, UnicodeToBinary(Chr(13)))

   If lngPosEnd > 0 Then
     strBinBoundary = MidB(strBinRequest, lngPosBeg, lngPosEnd - lngPosBeg)
     lngPosBoundary = InstrB(1, strBinRequest, strBinBoundary)
   End If
   If strBinBoundary = "" Then
     '-- Form is submitted without ENCTYPE="multipart/form-data" --
     lngPosBeg = 1
     lngPosEnd = InstrB(lngPosBeg, strBinRequest, UnicodeToBinary("&"))
     if lngPosEnd = 0 Then
       lngPosEnd = Len(strBinRequest) + 1
     End If
     While lngPosBeg < LenB(strBinRequest)
       strParameter = BinaryToUnicode(MidB(strBinRequest, lngPosBeg, _
                                           lngPosEnd - lbgPosBeg))
       lngPosDelimiter = Instr(1, strParameter, "=")
       strName = URLDecode(Left(strParameter, lngPosDelimiter - 1))
       strValue = URLDecode(Right(strParameter, _
                            Len(strParameter) - lngPosDelimiter))
       m_objForm.Add strName, strValue

       lngPosBeg = lngPosEnd + 1
       lngPosEnd = InstrB(lngPosBeg, strBinRequest, UnicodeToBinary("&"))
       If lngPosEnd = 0 Then
          lngPosEnd = Len(strBinRequest) + 1
       End If
     Wend
   Else
     '-- Form is submitted with ENCTYPE="multipart/form-data" --
     While lngPosBoundary < InstrB(strBinRequest, strBinBoundary & _
                                    UnicodeToBinary("--"))
       lngHelpPos = InstrB(lngPosBoundary, strBinRequest, UnicodeToBinary("Content-Disposition"))
       lngHelpPos = InstrB(lngHelpPos,strBinRequest, UnicodeToBinary("name="))

       lngPosBeg = lngHelpPos + 6
       lngPosEnd = InstrB(lngPosBeg, strBinRequest, UnicodeToBinary(Chr(34)))
       strName = BinaryToUnicode(MidB(strBinRequest, lngPosBeg, _
                                 lngPosEnd - lngPosBeg))
       '--- Get the file name ---
       lngPosFilename= InstrB(lngPosBoundary, strBinRequest, _
                              UnicodeToBinary("filename="))
       if (lngPosFileName <> 0) And _
          (lngPosFileName < InstrB(lngPosEnd, strBinRequest, strBinBoundary)) Then
         lngPosBeg = lngPosFileName + 10
         lngPosEnd = InstrB(lngPosBeg, strBinRequest, UnicodeToBinary(Chr(34)))
         strFileName = BinaryToUnicode(MidB(strBinRequest, lngPosBeg, _
                                       lngPosEnd - lngPosBeg))
         '-- Get the content type ---
         lngPosBeg = InstrB(lngPosEnd, strBinRequest, _
                            UnicodeToBinary("Content-Type:")) + 14
         lngPosEnd = InstrB(lngPosBeg, strBinRequest, UnicodeToBinary(Chr(13)))
         strContentType = BinaryToUnicode(MidB(strBinRequest, lngPosBeg, _
                                          lngPosEnd - lngPosBeg))
         '--- Get the content ---
         lngPosBeg = lngPosEnd + 4
         lngPosEnd = InstrB(lngPosBeg, strBinRequest, strBinBoundary) - 2
         strBinContent = MidB(strBinRequest, lngPosBeg, lngPosEnd - lngPosBeg)
         If (strFileName <> "") And (strBinContent <> "") Then
           '--- Create File object and Add to collection ---
           Set objFile = new clsFile
           objFile.Name = strName
           objFile.FileName = Right(strFileName, Len(strFileName) - _
                                    InStrRev(strFileName, "\"))
           objFile.ContentType = strContentType
           objFile.Blob = strBinContent
           m_objFiles.Add strName, objFile
         End If
       Else
         '--- Normal Form element ---
         lngHelpPos = InstrB(lngHelpPos, strBinRequest, UnicodeToBinary(Chr(13)))
         lngPosBeg = lngHelpPos + 4
         lngPosEnd = InstrB(lngPosBeg, strBinRequest, strBinBoundary) -2
         strValue = BinaryToUnicode(MidB(strBinRequest, lngPosBeg, _
                                    lngPosEnd - lngPosBeg))
         '-- Add element to collection ---
         m_objForm.Add strName, strValue
       End If
       '--- Move to next element ---
       lngPosBoundary = InstrB(lngPosBoundary + LenB(strBinBoundary), strBinRequest, _
                               strBinBoundary)
     Wend
   End If
 End Sub
End Class
%>