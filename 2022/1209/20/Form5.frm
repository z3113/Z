VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "2位数相加"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4680
   DrawMode        =   1  'Blackness
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form5.frx":0000
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "批改"
      Height          =   495
      Left            =   3000
      TabIndex        =   7
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "出题"
      Height          =   495
      Left            =   3000
      TabIndex        =   6
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   1080
      TabIndex        =   5
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   4335
   End
   Begin VB.Label Label3 
      Caption         =   "请填写答案："
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "加数2："
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "加数1："
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a%, b%
Option Explicit

Private Sub Command1_Click()
    Randomize
    a = Int(Rnd * 90 + 10)
    b = Int(Rnd * 90 + 10)
    Text1.Text = a
    Text2.Text = b
End Sub

Private Sub Command2_Click()
    If Val(Text3.Text) = 0 Then
        Label4.Caption = "请在结果文本框中输入计算结果"
    Else
        If a + b = Val(Text3.Text) Then
            Label4.Caption = "非常棒，再做下一题吧！"
        Else
            Label4.Caption = "答案错误，再试一次吧！"
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
