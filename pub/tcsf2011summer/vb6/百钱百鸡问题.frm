VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "����"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10635
   BeginProperty Font 
      Name            =   "����_GB2312"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   10635
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "�����Ǯ�ټ�"
      Default         =   -1  'True
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   4920
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   8640
      TabIndex        =   5
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   5325
      TabIndex        =   4
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2010
      TabIndex        =   3
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "С����"
      Height          =   495
      Left            =   6990
      TabIndex        =   2
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "ĸ����"
      Height          =   495
      Left            =   3675
      TabIndex        =   1
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "������"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   4080
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b, c As Single, see As Integer
Private Sub Command1_Click()
Cls
see = 0
a = CSng(Text1)
b = CSng(Text2)
c = CSng(Text3)
    For i = 1 To 100
        For j = 1 To 100 - i
            For k = 1 To 100 - i - j
                If i + j + k = 100 And i * a + j * b + k * c = 100 Then
                        Print "������"; i; "ֻ", "ĸ����"; j; "ֻ", "С����"; k; "ֻ"
                        see = see + 1
                    End If
            Next k
        Next j
    Next i
Print see, "��ѡ��"
End Sub
