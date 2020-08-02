<%
'--------------------------------------------------------------------------------
'- File        : clsFile.inc
'- Author      : H.J.N. Mayer
'- Description : Container class for file sent via an http multipart/form-data
'-               POST
'--------------------------------------------------------------------------------
Class clsFile

 Private m_strName
 Private m_strContentType
 Private m_strFileName
 Private m_vntBlob

 Private Sub Class_Initialize()
   m_strName = ""
   m_strContentType = ""
   m_strFileName = ""
   m_vntBlob = ""
 End Sub

 '-------------------------------------------------------------------------------
 ' Property : Name (Get)
 '-------------------------------------------------------------------------------
 Public Property Get Name()
   Name = m_strName
 End Property

 '-------------------------------------------------------------------------------
 ' Property : Name (Let)
 '-------------------------------------------------------------------------------
 Property Let Name(strName)
   m_strName = strName
 End Property

 '-------------------------------------------------------------------------------
 ' Property : ContentType (Get)
 '-------------------------------------------------------------------------------
 Property Get ContentType()
   ContentType = m_strContentType
 End Property

 '-------------------------------------------------------------------------------
 ' Property : ContentType (Let)
 '-------------------------------------------------------------------------------
 Property Let ContentType(strContentType)
   m_strContentType = strContentType
 End Property

 '-------------------------------------------------------------------------------
 ' Property : FileName (Get)
 '-------------------------------------------------------------------------------
 Public Property Get FileName()
   FileName = m_strFileName
 End Property

 '-------------------------------------------------------------------------------
 ' Property : FileName (Let)
 '-------------------------------------------------------------------------------
 Property Let FileName(strFileName)
   m_strFileName = strFileName
 End Property

 '-------------------------------------------------------------------------------
 ' Property : Blob (Get)
 '-------------------------------------------------------------------------------
 Public Property Get Blob()
   Blob = m_vntBlob
 End Property

 '-------------------------------------------------------------------------------
 ' Property : Blob (Let)
 '-------------------------------------------------------------------------------
 Property Let Blob(vntBlob)
   m_vntBlob = vntBlob
 End Property

 '-------------------------------------------------------------------------------
 ' Function : Save
 '-------------------------------------------------------------------------------
 Public Sub Save(strPath)
 Dim objFSO
 Dim objFile
 Dim lngCounter
   Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
   Set objFile = objFSO.CreateTextFile(strPath & "\" & m_strFileName)
   For lngCounter = 1 To LenB(m_vntBlob)
     objFile.Write Chr(AscB(MidB(m_vntBlob, lngCounter, 1)))
   Next
   objFile.Close
 End Sub
End Class
%>