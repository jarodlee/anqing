VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "单击计算"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3525
   BeginProperty Font 
      Name            =   "创艺简楷体"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   3525
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
    For i = 1 To 1000
        If i Mod 8 = 0 And i Mod 9 = 0 Then
            Print i
            s = s + 1
            a = a + i
        End If
    Next
        Print
        Print s; "个"
        Print "和："; a
End Sub

