VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   Picture         =   "计算三角形面积.frx":0000
   ScaleHeight     =   5670
   ScaleWidth      =   9060
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FFFF&
      Caption         =   "清除"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5520
      MaskColor       =   &H0000FFFF&
      TabIndex        =   5
      Top             =   4560
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "计算"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   990
      Left            =   5520
      TabIndex        =   3
      Top             =   720
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   21.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   855
      Left            =   1920
      TabIndex        =   2
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   21.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   855
      Left            =   1920
      TabIndex        =   1
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   21.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   735
      Left            =   1920
      TabIndex        =   0
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1815
      Left            =   5640
      TabIndex        =   4
      Top             =   2160
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b, c, s, area As Double



Private Sub Command1_Click()
    a = CDbl(Text1)
    b = CDbl(Text2)
    c = CDbl(Text3)
    s = (a + b + c) / 2
    area = Sqr(s * (s - a) * (s - b) * (s - c))
    Label1.Caption = "三角形面积：" + CStr(area)
End Sub

Private Sub Command2_Click()
    Text1 = ""
    Text2 = ""
    Text3 = ""
Label1.Caption = ""
End Sub
