VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   9510
   StartUpPosition =   3  '窗口缺省
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "文件浏览.frx":0000
      Left            =   4680
      List            =   "文件浏览.frx":0007
      TabIndex        =   3
      Text            =   "*.*"
      Top             =   4440
      Width           =   2655
   End
   Begin VB.FileListBox File1 
      Height          =   2610
      Left            =   4680
      TabIndex        =   2
      Top             =   1680
      Width           =   2655
   End
   Begin VB.DirListBox Dir1 
      Height          =   2610
      Left            =   1320
      TabIndex        =   1
      Top             =   2160
      Width           =   3015
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   1320
      TabIndex        =   0
      Top             =   1680
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim choicedfile As String

Private Sub Combo1_Click()
    File1.Pattern = Combo1.Text
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub List1_Click()

End Sub
