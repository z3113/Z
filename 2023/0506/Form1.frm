VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0E0FF&
   Caption         =   "数组1上机练习"
   ClientHeight    =   3765
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   ScaleHeight     =   3765
   ScaleWidth      =   7815
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFC0C0&
      Caption         =   "退出&Q"
      Height          =   495
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FF80FF&
      Caption         =   "代码调试2"
      Height          =   495
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FF80FF&
      Caption         =   "举例2"
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FF80FF&
      Caption         =   "举例1"
      Height          =   495
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FF80FF&
      Caption         =   "输出格式"
      Height          =   495
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF80FF&
      Caption         =   "平均值"
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF80FF&
      Caption         =   "代码调试1"
      Height          =   495
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1440
      Width           =   1455
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
    Form1.Hide
    Form6.Show
End Sub

Private Sub Command6_Click()
    Form1.Hide
    Form7.Show
End Sub

Private Sub Command7_Click()
    Unload Form1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("是否退出？", vbYesNo + 32, "退出") = vbNo Then Cancel = True
End Sub
