VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13245
   LinkTopic       =   "Form1"
   ScaleHeight     =   9390
   ScaleWidth      =   13245
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "清空"
      Height          =   495
      Left            =   11400
      TabIndex        =   3
      Top             =   8400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "输出"
      Height          =   495
      Left            =   7720
      TabIndex        =   2
      Top             =   8400
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   4040
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   8400
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "行数"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   8400
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    a = CInt(Text1)
    b = a \ 2
    For i = 1 To b
        Print Tab(a); "*"
        Print
        
End Sub

Private Sub Command2_Click()
    Cls
End Sub

Private Sub Form_Load()

End Sub
