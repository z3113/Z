VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "穷举法的应用"
   ClientHeight    =   4365
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form4"
   ScaleHeight     =   4365
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command8 
      Caption         =   "完成退出"
      Height          =   735
      Left            =   2400
      TabIndex        =   7
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      Caption         =   "猜灯谜"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "走台阶"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "勾股定理"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "对错题"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "找特殊数"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "换钱"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "买鸡"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Form4
    Form5.Show
End Sub

Private Sub Command2_Click()
    Unload Form4
    Form6.Show
End Sub

Private Sub Command3_Click()
    Unload Form4
    Form7.Show
End Sub

Private Sub Command4_Click()
    Unload Form4
    Form8.Show
End Sub

Private Sub Command5_Click()
    Unload Form4
    Form9.Show
End Sub

Private Sub Command6_Click()
    Unload Form4
    Form10.Show
End Sub

Private Sub Command7_Click()
    Unload Form4
    Form11.Show
End Sub

Private Sub Command8_Click()
    Unload Form4
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
