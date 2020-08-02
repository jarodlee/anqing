<%

'--------------------------------------------------------------------------------
'- Function    : URLDecode
'- Description : Decodes an URL encode string
'--------------------------------------------------------------------------------
Public Function URLDecode( strExpression)
Dim strSource
Dim strResult
Dim strChar
Dim lngCounter

  'Replace the '+' char by spaces
  strSource = Replace( Expression, "+", " ")

  'Replace Hex values for characters
  lngCounter = 0
  while (lngCounter < Len(strSource))
    lngCounter = lngCounter + 1
    strChar = Mid(strSource, lngCounter, 1)
    If strChar = "%" Then
      strChar = Chr(CInt("&H" & Mid(strSource, lngCounter+1, 2)))
      lngCounter = lngCounter + 2
    End If
    strResult = strResult & strChar
  wend

  URLDecode = strResult
End Function
%>