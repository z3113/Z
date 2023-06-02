VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0FF&
   Caption         =   "选择排序算法"
   ClientHeight    =   3975
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   7425
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "退出&Q"
      Height          =   495
      Left            =   5124
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2208
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "任务四&D"
      Height          =   495
      Left            =   3276
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2748
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "任务三&C"
      Height          =   495
      Left            =   1356
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "任务二&B"
      Height          =   495
      Left            =   3264
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1524
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "任务一&A"
      Height          =   495
      Left            =   1356
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "数组应用之选择排序算法"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   7455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Form1.Hide
    Form2.Show
End Sub

Private Sub Command2_Click()
    Form1.Hide
    Form3.Show
End Sub

Private Sub Command3_Click()
    Form1.Hide
    Form4.Show
End Sub

Private Sub Command4_Click()
    Form1.Hide
    Form5.Show
End Sub

Private Sub Command5_Click()
    Unload Form1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("确定要退出吗？", vbOKCancel + 32, "退出") = vbCancel Then Cancel = True
End Sub

