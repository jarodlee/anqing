VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15090
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   15090
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "华康简魏碑"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   8
      Text            =   "任意四则计算器.frx":0000
      Top             =   4080
      Width           =   14415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "华康简魏碑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10440
      TabIndex        =   7
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "计算"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "华康简魏碑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6720
      TabIndex        =   6
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "清除"
      BeginProperty Font 
         Name            =   "华康简魏碑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   5
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12000
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6900
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3630
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   36
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10200
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b, c  As Double, fuhao As String

Private Sub Command1_Click()
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Text4 = ""
End Sub

Private Sub Command2_Click()
    a = CDbl(Text1)
    b = CDbl(Text3)
    fuhao = Text2
If Not (IsNumeric(a) Or IsNumeric(b)) Then
    MsgBox ("输入不合法，请检查！")
        ElseIf fuhao = "+" Then
            c = a + b
                ElseIf fuhao = "-" Then
                    c = a - b
                        ElseIf fuhao = "*" Then
                            c = a * b
                                ElseIf fuhao = "/" Then
                                    c = a / b
                                        Else
                                            MsgBox ("输入不合法，请检查！")
        End If
        Text4 = CStr(c)
End Sub

Private Sub Command3_Click()
    End
End Sub

Private Sub Form_Load()
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Text4 = ""
End Sub

