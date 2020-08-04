VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000D&
   Caption         =   "Form1"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11595
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   11595
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "清除"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      Height          =   735
      Left            =   4980
      TabIndex        =   5
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "华转摄"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8640
      TabIndex        =   4
      Top             =   3720
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "摄转华"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      TabIndex        =   3
      Top             =   3720
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6720
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1560
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   720
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1560
      Width           =   4095
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "华氏温度"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   975
      Left            =   7680
      TabIndex        =   6
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "摄氏温度"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b, c As Double

Private Sub Command1_Click()
    a = CDbl(Text1)
    c = 9 * a / 5 + 32
    Text2 = CStr(c)
End Sub

Private Sub Command2_Click()

    b = CDbl(Text2)
    c = 5 * (b - 32) / 9
    Text1 = CStr(c)
End Sub

Private Sub Command3_Click()
    Text1 = ""
    Text2 = ""
End Sub

Private Sub Form_Load()
    Text1 = ""
    Text2 = ""
    a = 0
    b = 0
End Sub


