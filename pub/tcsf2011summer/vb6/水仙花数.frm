VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "水仙花数"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   7230
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1440
      Width           =   4695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "输出"
      BeginProperty Font 
         Name            =   "华康简魏碑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "1000以内的水仙花数"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   480
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    For i = 100 To 999
        a = i \ 100
        b = (i - a * 100) \ 10
        c = i Mod 10
            If i = a ^ 3 + b ^ 3 + c ^ 3 Then
                Text1.Text = Text1.Text + CStr(i) + vbCrLf
            End If
    Next i
End Sub
