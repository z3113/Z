VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "评委打分"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   6045
   StartUpPosition =   3  '窗口缺省
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
      Left            =   4080
      TabIndex        =   8
      Top             =   3480
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
      Left            =   2400
      TabIndex        =   7
      Top             =   3480
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
      Left            =   720
      TabIndex        =   6
      Top             =   3480
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
      Index           =   2
      Left            =   4080
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
      Index           =   1
      Left            =   2400
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "高到低排序"
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
      Left            =   4080
      TabIndex        =   3
      Top             =   4548
      Width           =   1455
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
      Left            =   720
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
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
      Left            =   480
      TabIndex        =   1
      Top             =   4548
      Width           =   1335
   End
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
      Left            =   2280
      TabIndex        =   0
      Top             =   4548
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   5
      Left            =   4080
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   4
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   3
      Left            =   720
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   2
      Left            =   4080
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   1
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   975
      Index           =   0
      Left            =   720
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a(6) As Integer

Private Sub Command1_Click()
    Dim i%
    For i = 0 To 5
        a(i) = Int(Rnd * 11)
        Text1(i).Text = a(i)
    Next i
End Sub

Private Sub Command2_Click()
    Dim i%, b#, min%, max%
    min = a(0)
    max = a(0)
    For i = 0 To 5
        If min > a(i) Then min = a(i)
        If max < a(i) Then max = a(i)
        b = b + a(i)
    Next i
    b = Round((b - min - max) / 4)
    MsgBox "去掉一个最高分和一个最低分，该选手的平均得分是" & b, vbOKOnly, "最终结果"
End Sub

Private Sub Command3_Click()
    Dim i%, j%, b%, c(6) As String, d%
    For i = 0 To 5
        c(i) = i
    Next i
    For i = 0 To 4
        b = i
        For j = i + 1 To 5
            If a(b) < a(j) Then b = j
        Next j
        If i <> b Then a(6) = a(i): a(i) = a(b): a(b) = a(6): c(6) = c(i): c(i) = c(b): c(b) = c(6)
    Next i
    For i = 0 To 5
        Image1(i).Picture = LoadPicture(c(i) & ".JPG")
        Text1(i).Text = a(i)
    Next i
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

Private Sub Text1_Change(Index As Integer)
    Text1(Index).ToolTipText = "这是第" & Index + 1 & "个文本框"
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    Text1(Index).BackColor = vbYellow
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).BackColor = &H80000005
End Sub
