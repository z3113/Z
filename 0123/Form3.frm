VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "简单应用"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3720
      Top             =   1080
   End
   Begin VB.CommandButton Command1 
      Height          =   735
      Left            =   1440
      Picture         =   "Form3.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   1
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "班级幸运儿产生："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a%

Private Sub Command1_Click()
    If a = 0 Then
        Timer1.Enabled = True
        Label1.Caption = "抽取幸运儿："
        a = 1
    ElseIf a = 1 Then
        Timer1.Enabled = False
        Label1.Caption = "班级的幸运儿是学号：" & Text1.Text
        a = 0
    End If
End Sub

Private Sub Form_Activate()
    a = 0
End Sub

Private Sub Form_DblClick()
    Unload Form3
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub

Private Sub Timer1_Timer()
    Randomize
    Text1.Text = Int(Rnd * 34 + 1) & "号"
End Sub
