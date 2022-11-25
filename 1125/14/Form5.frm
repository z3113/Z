VERSION 5.00
Begin VB.Form Form5 
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8160
   LinkTopic       =   "Form5"
   ScaleHeight     =   5355
   ScaleWidth      =   8160
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "开始抽取学号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1560
      TabIndex        =   4
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "停止"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5400
      TabIndex        =   3
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   3000
      TabIndex        =   2
      Text            =   "40"
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   5760
      Top             =   1560
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "返回"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6720
      TabIndex        =   1
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      TabIndex        =   0
      Text            =   "40"
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "本班级人数"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   5
      Top             =   1560
      Width           =   2535
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a%

Private Sub Command1_Click()
    Timer1.Enabled = False
End Sub

Private Sub Command2_Click()
    Unload Form5
    Form4.Show
End Sub

Private Sub Command3_Click()
    a = Val(Text1.Text)
    Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Form5
    Form4.Show
End Sub

Private Sub Timer1_Timer()
    Randomize
    Text2.Text = Int(Rnd * a + 1)
End Sub
