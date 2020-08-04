VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5925
   ClientLeft      =   3900
   ClientTop       =   3195
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   10965
   Begin VB.TextBox Text1 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2175
      Left            =   1080
      TabIndex        =   2
      Top             =   600
      Width           =   9255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   495
      Left            =   9480
      TabIndex        =   1
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton 显示图片 
      Caption         =   "显示图片"
      Height          =   495
      Left            =   7800
      TabIndex        =   0
      Top             =   5160
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Text1_Click()
Text1.Text = "这是一个窗体操作的实例。"
End Sub

Private Sub 显示图片_Click()
Form1.Picture = LoadPicture("C:\WINDOWS\Web\Wallpaper\1.jpg")

End Sub
