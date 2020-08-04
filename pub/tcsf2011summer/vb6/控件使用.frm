VERSION 5.00
Begin VB.Form 定时器的使用 
   Caption         =   "Form1"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12720
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   12720
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "继续"
      Height          =   495
      Left            =   10320
      TabIndex        =   2
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "停止"
      Height          =   495
      Left            =   8160
      TabIndex        =   1
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Interval        =   300
      Left            =   5760
      Top             =   4200
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   5400
      Top             =   240
   End
   Begin VB.Image Image1 
      Height          =   3375
      Left            =   720
      Picture         =   "控件使用.frx":0000
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Image Image2 
      Height          =   3375
      Left            =   7920
      Picture         =   "控件使用.frx":19F85
      Stretch         =   -1  'True
      Top             =   1320
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "全家幸福"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   4440
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "定时器的使用"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Picture1_Click()

End Sub

Private Sub Command1_Click()
    Timer1.Enabled = False
    Timer2.Enabled = False
End Sub

Private Sub Command2_Click()
    Timer1.Enabled = True
    Timer2.Enabled = True
End Sub

Private Sub Timer1_Timer()
Label1.Left = Label1.Left + 20
End Sub

Private Sub Timer2_Timer()
Image1.Visible = Not (Image1.Visible)
Image2.Visible = Not (Image2.Visible)
End Sub
