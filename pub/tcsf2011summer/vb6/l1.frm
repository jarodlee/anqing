VERSION 5.00
Begin VB.Form ��ӭ 
   Caption         =   "��ӭ"
   ClientHeight    =   4695
   ClientLeft      =   6060
   ClientTop       =   5925
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   7950
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   1455
      Left            =   6840
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton ��ʾ 
      Caption         =   "��ʾ"
      Height          =   855
      Left            =   2640
      TabIndex        =   1
      Top             =   2400
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
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
Attribute VB_Name = "��ӭ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Text1.Text = ""
    
End Sub


Private Sub Text1_Change()

End Sub

Private Sub ��ʾ_Click()
    Text1.Text = "������.������=����ֵ"
End Sub
