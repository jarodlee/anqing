VERSION 5.00
Begin VB.MDIForm mainfrm 
   BackColor       =   &H8000000C&
   Caption         =   "多文档编辑器"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  '窗口缺省
   Begin VB.Menu mnufile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnunew 
         Caption         =   "新建(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnubar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "退出(&Q)"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "编辑(&e)"
      Begin VB.Menu mnucut 
         Caption         =   "剪切(&x)"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuwindows 
      Caption         =   "窗口(&w)"
      Begin VB.Menu mnuhor 
         Caption         =   "水平平铺(&t)"
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

