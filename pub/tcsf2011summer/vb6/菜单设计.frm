VERSION 5.00
Begin VB.MDIForm mainfrm 
   BackColor       =   &H8000000C&
   Caption         =   "���ĵ��༭��"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Menu mnufile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnunew 
         Caption         =   "�½�(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnubar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "�˳�(&Q)"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "�༭(&e)"
      Begin VB.Menu mnucut 
         Caption         =   "����(&x)"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuwindows 
      Caption         =   "����(&w)"
      Begin VB.Menu mnuhor 
         Caption         =   "ˮƽƽ��(&t)"
      End
   End
End
Attribute VB_Name = "mainfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim newform As childfrm

Private Sub MDIForm_Load()
    childfrm.Hide
    mainfrm.WindowState = vbMaximized
End Sub

