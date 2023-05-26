VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "数组、排序综合"
   ClientHeight    =   4470
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   8175
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "标签下落"
      Height          =   495
      Index           =   1
      Left            =   1800
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "冒泡排序"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "评委打分"
      Height          =   495
      Index           =   3
      Left            =   5160
      TabIndex        =   2
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "求素数"
      Height          =   495
      Index           =   2
      Left            =   3480
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "计算器"
      Height          =   495
      Index           =   4
      Left            =   6840
      TabIndex        =   0
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   " "
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   7200
      TabIndex        =   5
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n%
Private Sub Command1_Click(Index As Integer)
    Me.Hide
    Select Case Index
        Case 0: Form2.Show
        Case 1: Form3.Show
        Case 2: Form4.Show
        Case 3: Form5.Show
        Case 4: Form6.Show
    End Select
End Sub

Private Sub Form_Load()
    n = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MsgBox "你用了" & Label2.Caption & "秒", vbYesNo + 64, "提示"
    If MsgBox("你确定要退出吗？", vbYesNo + 32, "提示") = vbNo Then Cancel = True
End Sub

Private Sub Timer1_Timer()
    Label2.Caption = n
    n = n + 1
End Sub
