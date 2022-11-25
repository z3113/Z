VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   4470
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   5700
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFC0C0&
      Caption         =   "开始"
      Height          =   975
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Cancel          =   -1  'True
      Caption         =   "结束(ESC)"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "任务五(&E)"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "任务四(&D)"
      Enabled         =   0   'False
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "任务三(&C)"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "任务二(&B)"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "任务一（&A)"
      Enabled         =   0   'False
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "单击窗体改变背景色"
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   5535
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
    End
End Sub

Private Sub Command7_Click()
    Command1.Enabled = True
    Command2.Enabled = True
    Command3.Enabled = True
    Command4.Enabled = True
    Command5.Enabled = True
    Command6.Enabled = True
End Sub

Private Sub Form_Click()
    Randomize
    Form1.BackColor = RGB(Int(Rnd * 256), Int(Rnd * 256), Int(Rnd * 256))
End Sub

Private Sub Label1_Click()

End Sub
