VERSION 5.00
Begin VB.Form 按钮的使用 
   Caption         =   "按钮的使用"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   8190
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "退出"
      Height          =   495
      Left            =   5520
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清除"
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "显示"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   1455
      Left            =   2160
      TabIndex        =   0
      Top             =   1080
      Width           =   4095
   End
End
Attribute VB_Name = "按钮的使用"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Text1.Text = "这玩意儿不错！"
End Sub

Private Sub Command2_Click()
    Text1.Text = ""
End Sub

Private Sub Command3_Click()
End
End Sub
