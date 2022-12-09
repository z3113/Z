VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0FFC0&
   Caption         =   "IF语句双分支练习"
   ClientHeight    =   3390
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   ScaleHeight     =   3390
   ScaleWidth      =   5205
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command7 
      Caption         =   "结束"
      Height          =   495
      Left            =   2040
      TabIndex        =   6
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "任务六"
      Height          =   495
      Left            =   3480
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "任务五"
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "任务四"
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "任务三"
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "任务二"
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "任务一"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "IF双分支语句练习"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   840
      TabIndex        =   7
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Form1
    Form2.Show
End Sub

Private Sub Command2_Click()
    Unload Form1
    Form3.Show
End Sub

Private Sub Command3_Click()
    Unload Form1
    Form4.Show
End Sub

Private Sub Command4_Click()
    Unload Form1
    Form5.Show
End Sub

Private Sub Command5_Click()
    Unload Form1
    Form6.Show
End Sub

Private Sub Command6_Click()
    Unload Form1
    Form7.Show
End Sub

Private Sub Command7_Click()
    End
End Sub
