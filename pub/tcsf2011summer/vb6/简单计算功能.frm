VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   11850
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "退出"
      Height          =   735
      Left            =   9960
      TabIndex        =   7
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清除"
      Height          =   855
      Left            =   1200
      TabIndex        =   6
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "运算"
      Height          =   855
      Left            =   4680
      TabIndex        =   5
      Top             =   3480
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   9000
      TabIndex        =   4
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4800
      TabIndex        =   2
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7680
      TabIndex        =   3
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer, b As Integer, c As Integer



Private Sub Command1_Click()
a = CInt(Text1)
b = CInt(Text2)
c = a + b
Text3 = CStr(c)
End Sub

Private Sub Command2_Click()
Text1 = ""
Text2 = ""
Text3 = ""
End Sub

Private Sub Command3_Click()
End
End Sub
