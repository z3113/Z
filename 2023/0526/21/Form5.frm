VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "评委打分"
   ClientHeight    =   5535
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6615
   LinkTopic       =   "Form5"
   ScaleHeight     =   5535
   ScaleWidth      =   6615
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "平均得分"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "评委打分"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   4560
      TabIndex        =   7
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   2760
      TabIndex        =   6
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   960
      TabIndex        =   5
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   960
      TabIndex        =   4
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   2760
      TabIndex        =   3
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   4560
      TabIndex        =   2
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   0
      Left            =   960
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   1
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   2
      Left            =   4560
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   3
      Left            =   960
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   4
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   5
      Left            =   4560
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1095
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(6) As Integer

Private Sub Command1_Click()
    Dim i%
    For i = 1 To 6
        a(i) = Int(Rnd * 11)
        Text1(i - 1).Text = a(i)
    Next i
End Sub

Private Sub Command2_Click()
    Dim i%, b#, min%, max%
    min = a(1)
    max = a(1)
    For i = 1 To 6
        b = b + a(i)
        If min > a(i) Then min = a(i)
        If max < a(i) Then max = a(i)
    Next i
    b = Round((b - min - max) / 6, 2)
    MsgBox "去掉一个最高分和一个最低分，该选手的平均得分是" & b, vbOKOnly, "最终结果"
End Sub

Private Sub Form_Activate()
    Dim i%
    For i = 0 To 5
        Image1(i).Picture = LoadPicture(i & ".JPG")
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    Text1(Index).BackColor = vbYellow
    Text1(Index).ToolTipText = "这是第" & Index + 1 & "个文本框"
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).BackColor = &H80000005
End Sub
