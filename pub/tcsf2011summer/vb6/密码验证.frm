VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "用户登录"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   4845
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "登录"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3720
      TabIndex        =   5
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "$"
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "华康简魏碑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   2520
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "密码"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "用户名"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim username, passwd As String

Private Sub Command1_Click()
    If Trim(Text1.Text) = username And Trim(Text2.Text) = passwd Then
        Label3.Caption = "登录成功”"
        MsgBox ("欢迎光临GHTxx.cn")
        Else
            Label3.Caption = "登录失败"
            MsgBox ("别试了，找城管抓你^_^")
            Text1 = ""
            Text2 = ""
    End If
End Sub

Private Sub Form_Load()
    username = "Jarod"
    passwd = "123"
End Sub
