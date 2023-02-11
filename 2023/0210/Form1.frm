VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "vb综合复习"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   7095
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command8 
      Caption         =   "退出"
      Height          =   615
      Left            =   5520
      TabIndex        =   8
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      Caption         =   "移动文字"
      Height          =   615
      Left            =   3720
      TabIndex        =   7
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "税金计算"
      Height          =   615
      Left            =   1920
      TabIndex        =   6
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "综合计算"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "最值"
      Height          =   615
      Left            =   5520
      TabIndex        =   4
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "素数"
      Height          =   615
      Left            =   3720
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "字体字号"
      Height          =   615
      Left            =   1920
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "倒计时"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "VB综合复习"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
    Form1.Hide
    Form8.Show
End Sub

Private Sub Command8_Click()
    Unload Form1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("确定关闭窗口吗？", vbOKCancel + 16, "关闭窗口") = vbCancel Then Cancel = True
End Sub
