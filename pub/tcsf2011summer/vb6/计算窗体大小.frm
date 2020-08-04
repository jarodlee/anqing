VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton exit 
      Caption         =   "退出"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton zoom 
      Caption         =   "窗体变大"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub exit_Click()
End
End Sub

Private Sub Form_Click()
  Form1.Width = Form1.Width - 100
    Form1.Height = Form1.Height - 100
End Sub

Private Sub zoom_Click()
    Form1.Width = Form1.Width + 100
    Form1.Height = Form1.Height + 100
    
End Sub
