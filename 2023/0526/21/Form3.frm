VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "标签下落"
   ClientHeight    =   7050
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12960
   LinkTopic       =   "Form3"
   ScaleHeight     =   7050
   ScaleWidth      =   12960
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   12120
      Top             =   840
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      Height          =   495
      Left            =   11760
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Height          =   615
      Index           =   13
      Left            =   10920
      TabIndex        =   14
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Height          =   615
      Index           =   12
      Left            =   10080
      TabIndex        =   13
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Height          =   615
      Index           =   11
      Left            =   9240
      TabIndex        =   12
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Height          =   615
      Index           =   10
      Left            =   8400
      TabIndex        =   11
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Height          =   615
      Index           =   9
      Left            =   7560
      TabIndex        =   10
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Height          =   615
      Index           =   8
      Left            =   6720
      TabIndex        =   9
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Height          =   615
      Index           =   7
      Left            =   5880
      TabIndex        =   8
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Height          =   615
      Index           =   6
      Left            =   5040
      TabIndex        =   7
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Height          =   615
      Index           =   5
      Left            =   4200
      TabIndex        =   6
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Height          =   615
      Index           =   4
      Left            =   3360
      TabIndex        =   5
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Height          =   615
      Index           =   3
      Left            =   2520
      TabIndex        =   4
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Height          =   615
      Index           =   2
      Left            =   1680
      TabIndex        =   3
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Height          =   615
      Index           =   1
      Left            =   840
      TabIndex        =   2
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Height          =   615
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(13) As Integer
Private Sub Command1_Click()
    Timer1.Enabled = True
    For i = 0 To 13
        a(i) = Rnd * 41 + 10
    Next i
End Sub

Private Sub Form_Activate()
    Randomize
    Dim i%
    For i = 0 To 13
        Label1(i).BackColor = RGB(Int(Rnd * 256), Int(Rnd * 256), Int(Rnd * 256))
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub

Private Sub Timer1_Timer()
    Randomize
    For i = 0 To 13
        Label1(i).Top = Label1(i).Top + a(i)
    Next i
End Sub
