VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "home"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9030
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0B3A
   ScaleHeight     =   4455
   ScaleWidth      =   9030
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command8 
      Caption         =   "退出"
      Height          =   615
      Left            =   6360
      TabIndex        =   7
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      Caption         =   "提高"
      Height          =   615
      Left            =   4440
      TabIndex        =   6
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "分段函数"
      Height          =   615
      Left            =   2520
      TabIndex        =   5
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "数位处理"
      Height          =   615
      Left            =   600
      TabIndex        =   4
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "征  税"
      Height          =   615
      Left            =   6360
      TabIndex        =   3
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "商场打折"
      Height          =   615
      Left            =   4440
      TabIndex        =   2
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "成绩处理"
      Height          =   615
      Left            =   2520
      TabIndex        =   1
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "符号函数"
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ELSEIF语句练习"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1920
      TabIndex        =   8
      Top             =   480
      Width           =   3975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Form2.Hide
    Form3.Show
End Sub

Private Sub Command2_Click()
    Form2.Hide
    Form4.Show
End Sub

Private Sub Command3_Click()
    Form2.Hide
    Form5.Show
End Sub

Private Sub Command4_Click()
    Form2.Hide
    Form6.Show
End Sub

Private Sub Command5_Click()
    Form2.Hide
    Form7.Show
End Sub

Private Sub Command6_Click()
    Form2.Hide
    Form8.Show
End Sub

Private Sub Command7_Click()
    Form2.Hide
    Form9.Show
End Sub

Private Sub Command8_Click()
    End
End Sub
