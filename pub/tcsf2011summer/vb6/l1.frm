VERSION 5.00
Begin VB.Form 欢迎 
   Caption         =   "欢迎"
   ClientHeight    =   4695
   ClientLeft      =   6060
   ClientTop       =   5925
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   7950
   Begin VB.CommandButton Command1 
      Caption         =   "隐藏"
      Height          =   1455
      Left            =   6840
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton 显示 
      Caption         =   "显示"
      Height          =   855
      Left            =   2640
      TabIndex        =   1
      Top             =   2400
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1455
      Left            =   1200
      TabIndex        =   0
      Top             =   600
      Width           =   5415
   End
End
Attribute VB_Name = "欢迎"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Text1.Text = ""
    
End Sub


Private Sub Text1_Change()

End Sub

Private Sub 显示_Click()
    Text1.Text = "对象名.属性名=属性值"
End Sub
