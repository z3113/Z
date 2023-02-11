VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   9270
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command6 
      Enabled         =   0   'False
      Height          =   495
      Left            =   7440
      TabIndex        =   5
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Enabled         =   0   'False
      Height          =   495
      Left            =   5640
      TabIndex        =   4
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Enabled         =   0   'False
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Enabled         =   0   'False
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Enabled         =   0   'False
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   135
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   135
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
    Command4.Enabled = True
    Command5.Enabled = True
    Command6.Enabled = True
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
    Form1.Hide
    Form4.Show
End Sub

Private Sub Command5_Click()
    Form1.Hide
    Form5.Show
End Sub

Private Sub Command6_Click()
    Form1.Hide
    Form6.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MsgBox "今天就学到这里吧，再见！", 1 + 64, "退出消息框"
End Sub

