VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10680
   BeginProperty Font 
      Name            =   "Romana BT"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9585
   ScaleWidth      =   10680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "阶乘"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   5520
      TabIndex        =   1
      Top             =   1560
      Width           =   3615
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
      Height          =   1095
      Left            =   5520
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label1 
      Height          =   1215
      Left            =   5520
      TabIndex        =   2
      Top             =   3120
      Width           =   4935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, s As Double

Private Sub Command1_Click()
    Cls
    a = CInt(Text1)
    s = 1
        For i = 1 To a Step 1
            s = i * s
            Print s
        Next i
    Label1.Caption = CStr(s)
End Sub
