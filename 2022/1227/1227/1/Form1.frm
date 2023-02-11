VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "高二VB期末考试"
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   3660
   ScaleWidth      =   8895
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "退出"
      Height          =   615
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "综合应用"
      Height          =   615
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "简单应用"
      Height          =   615
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "基本操作"
      Height          =   615
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "VB程序设计期末考试试题"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   30
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   0
      Top             =   1200
      Width           =   6975
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
    form3.Show
End Sub

Private Sub Command3_Click()
    Form1.Hide
    form4.Show
End Sub

Private Sub Command4_Click()
    Unload Form1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("是否退出", vbYesNo + 32, "退出") = vbNo Then Cancel = True
End Sub
