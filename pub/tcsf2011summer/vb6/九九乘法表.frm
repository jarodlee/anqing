VERSION 5.00
Begin VB.Form 九九乘法表 
   Caption         =   "Form1"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13680
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   13680
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1080
      Width           =   12975
   End
   Begin VB.CommandButton hello 
      BackColor       =   &H008080FF&
      Caption         =   "输出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "九九乘法表"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim shower As String, i, j, k As Integer
Private Sub hello_Click()
    For i = 1 To 9
        For j = 1 To i
            k = i * j
            If i < j Then
                shower = CStr(i) + "×" + CStr(j) + "=" + CStr(k)
            Else
                shower = CStr(j) + "×" + CStr(i) + "=" + CStr(k)
            End If
            If k < 10 Then
                Text1.Text = Text1.Text + shower + "       "
                    Else
                        Text1.Text = Text1.Text + shower + "     "
            End If
        Next j
            Text1.Text = Text1.Text + vbCrLf
    Next i
End Sub
