VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0E0FF&
   Caption         =   "分支结构综合练习"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8205
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   8205
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "退出（ESC）"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   7935
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "任务二&B"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "任务一&A"
      Enabled         =   0   'False
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始（ENTER）"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   7935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Command2.Enabled = True
    Command3.Enabled = True
End Sub

Private Sub Command2_Click()
    Form1.Hide
    Form2.Show
End Sub

Private Sub Command3_Click()
    Form1.Hide
    Form3.Show
End Sub

Private Sub Command4_Click()
    End
End Sub
