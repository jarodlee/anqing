VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12030
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   14.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3630
   ScaleWidth      =   12030
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "计算"
      Height          =   495
      Left            =   10440
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   8520
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   4680
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "增长率         　 ％"
      Height          =   495
      Left            =   7200
      TabIndex        =   5
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "预期数           亿"
      Height          =   495
      Left            =   3600
      TabIndex        =   4
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "人口数            亿"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
    i = 13
    Do Until i >= 26
    i = i + i * 0.008
    j = j + 1
    Loop
    Print "经过"; j; "年后，人口为"; i; "亿"
End Sub
