VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   ScaleHeight     =   9345
   ScaleWidth      =   10320
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmdexit 
      Caption         =   "�˳�"
      Height          =   495
      Left            =   8760
      TabIndex        =   17
      Top             =   8520
      Width           =   1215
   End
   Begin VB.CommandButton cmdlast 
      Caption         =   "���һ��"
      Height          =   495
      Left            =   8760
      TabIndex        =   16
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton cmdfirst 
      Caption         =   "��һ��"
      Height          =   495
      Left            =   7320
      TabIndex        =   15
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "��һ��"
      Height          =   495
      Left            =   8760
      TabIndex        =   14
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "��һ��"
      Height          =   495
      Left            =   7320
      TabIndex        =   13
      Top             =   6600
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      DataField       =   "�й�����"
      DataSource      =   "Data1"
      Height          =   1455
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   12
      Text            =   "ͼ�����ϲ�ѯ.frx":0000
      Top             =   7200
      Width           =   5415
   End
   Begin VB.TextBox Text5 
      DataField       =   "������"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   1800
      TabIndex        =   11
      Text            =   "Text5"
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      DataField       =   "����"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   1800
      TabIndex        =   10
      Text            =   "Text4"
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      DataField       =   "��Ŀ"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   1800
      TabIndex        =   9
      Text            =   "Text3"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      DataField       =   "����"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   1800
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataField       =   "���"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   1800
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   840
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "E:\www\vb6\book-list.mdb"
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
      Exclusive       =   0   'False
      Height          =   1575
      Left            =   8040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Bookform"
      Top             =   480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   8295
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   6255
      Begin VB.Label Label6 
         Caption         =   "�й�����"
         Height          =   735
         Left            =   480
         TabIndex        =   6
         Top             =   6120
         Width           =   5415
      End
      Begin VB.Label Label5 
         Caption         =   "������"
         Height          =   735
         Left            =   480
         TabIndex        =   5
         Top             =   4944
         Width           =   5415
      End
      Begin VB.Label Label4 
         Caption         =   "����"
         Height          =   735
         Left            =   480
         TabIndex        =   4
         Top             =   3768
         Width           =   5415
      End
      Begin VB.Label Label3 
         Caption         =   "��Ŀ"
         Height          =   735
         Left            =   480
         TabIndex        =   3
         Top             =   2592
         Width           =   5415
      End
      Begin VB.Label Label2 
         Caption         =   "����"
         Height          =   735
         Left            =   480
         TabIndex        =   2
         Top             =   1416
         Width           =   5415
      End
      Begin VB.Label Label1 
         Caption         =   "���"
         Height          =   735
         Left            =   480
         TabIndex        =   1
         Top             =   240
         Width           =   5415
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdexit_Click()
End
End Sub

Private Sub cmdfirst_Click()
    Data1.Recordset.MoveFirst
    cmdPrev.Enabled = False
    cmdNext.Enabled = True
End Sub

Private Sub cmdlast_Click()
    Data1.Recordset.MoveLast
    cmdPrev.Enabled = True
    cmdNext.Enabled = False
End Sub

Private Sub cmdNext_Click()
    Data1.Recordset.MoveNext
    cmdNext.Enabled = True
        If Data1.Recordset.EOF Then
             Data1.Recordset.MoveLast
             cmdPrev.Enabled = False
        End If
End Sub

Private Sub cmdPrev_Click()
    Data1.Recordset.MovePrevious
    cmdNext.Enabled = True
        If Data1.Recordset.BOF Then
            Data1.Recordset.MoveFirst
            cmdPrev.Enabled = False
        End If
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path + "\book-list.mdb"
Data1.RecordSource = "bookform"
End Sub
