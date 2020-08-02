<%

'--------------------------------------------------------------------------------
'- Function    : BinaryToUnicode
'- Description : Converts a Binary string to a Unicode string
'--------------------------------------------------------------------------------
Public Function BinaryToUnicode(strBinary)
Dim lngCounter
Dim strUnicode
  strUnicode = ""
  For lngCounter = 1 To LenB(strBinary)
    strUnicode = strUnicode & Chr(AscB(MidB(strBinary, lngCounter, 1)))
  Next
  BinaryToUnicode = strUnicode
End Function

'--------------------------------------------------------------------------------
'- Function    : UnicodeToBinary
'- Description : Converts a Unicode string to a Binary string
'--------------------------------------------------------------------------------
Public Function UnicodeToBinary(strUnicode)
Dim lngCounter
Dim strChar
Dim strBinary
  strBinary = ""
  For lngCounter = 1 To Len(strUnicode)
    strChar = Mid(strUnicode, lngCounter, 1)
    strBinary = strBinary & ChrB(AscB(strChar))
  Next
  UnicodeToBinary = strBinary
End Function
%>