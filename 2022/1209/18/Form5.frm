VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H0080C0FF&
   Caption         =   "回答判断"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6945
   LinkTopic       =   "Form5"
   ScaleHeight     =   4230
   ScaleWidth      =   6945
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5520
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "请输入y或n："
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "今天你高兴吗？"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   1440
      Width           =   2175
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim a$
    a = LCase(Text1.Text)
    If a <> "y" And a <> "n" Then
        Print "Please enter the y or n"
    Else
        If a = "y" Then
            Print "you are happy"
        Else
            Print "you are not happy"
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form2.Show
End Sub
