VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3555
   ScaleWidth      =   6000
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Caption         =   "ÆÀ¶¨"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "»ª¿µ¼òÎº±®"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      MaskColor       =   &H000000FF&
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   26.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "»ª¿µ¼òÎº±®"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   2640
      TabIndex        =   1
      Top             =   480
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Double

Private Sub Command1_Click()
    a = CDbl(Text1.Text)
    If a > 100 Then
        Label1.Caption = "Äã³­¶àÁË£¡"
        ElseIf a >= 85 Then
        Label1.Caption = "ÓÅÐã"
            ElseIf a >= 75 Then
                Label1.Caption = "Á¼"
                    ElseIf a >= 60 Then
                        Label1.Caption = "¼°¸ñ"
                            Else
                                Label1.Caption = "²î"
    End If
End Sub
