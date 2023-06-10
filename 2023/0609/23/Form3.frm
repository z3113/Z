VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FFC0C0&
   Caption         =   "简单应用"
   ClientHeight    =   6495
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7695
   LinkTopic       =   "Form3"
   ScaleHeight     =   6495
   ScaleWidth      =   7695
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "返回（&E)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5640
      TabIndex        =   6
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "排序"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3960
      TabIndex        =   5
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "显示"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      TabIndex        =   4
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "产生数据"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   3
      Top             =   4800
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   840
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   2
      Top             =   3600
      Width           =   6015
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3195
      ItemData        =   "Form3.frx":0000
      Left            =   4560
      List            =   "Form3.frx":0002
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3195
      ItemData        =   "Form3.frx":0004
      Left            =   840
      List            =   "Form3.frx":0006
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(50) As Integer
Private Sub Command1_Click()
    Randomize
    Dim i%
    List1.Clear
    For i = 1 To 50
        a(i) = Int(Rnd * 90 + 10)
        List1.AddItem a(i)
    Next i
End Sub

Private Sub Command2_Click()
    Dim i%
    Text1.Text = ""
    For i = 0 To 49
        Text1.Text = Text1.Text & List1.List(i) & "  "
    Next i
End Sub

Private Sub Command3_Click()
    Dim i%, j%
    List2.Clear
    For i = 1 To 49
        For j = 1 To 50 - i
            If a(j) > a(j + 1) Then a(0) = a(j): a(j) = a(j + 1): a(j + 1) = a(0)
        Next j
    Next i
    For i = 1 To 50
        List2.AddItem a(i)
    Next i
End Sub

Private Sub Command4_Click()
    Unload Form3
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
