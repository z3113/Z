VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00FFC0C0&
   Caption         =   "综合应用"
   ClientHeight    =   6510
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8775
   LinkTopic       =   "Form5"
   ScaleHeight     =   6510
   ScaleWidth      =   8775
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command6 
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
      Left            =   5760
      TabIndex        =   7
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "找素数"
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
      Left            =   3600
      TabIndex        =   6
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "找奇数"
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
      Left            =   1440
      TabIndex        =   5
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "找最大值"
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
      Left            =   5760
      TabIndex        =   4
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "求平均值"
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
      Left            =   3600
      TabIndex        =   3
      Top             =   4200
      Width           =   1575
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
      Left            =   1440
      TabIndex        =   2
      Top             =   4200
      Width           =   1575
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
      Height          =   2910
      ItemData        =   "Form5.frx":0000
      Left            =   5160
      List            =   "Form5.frx":0002
      TabIndex        =   1
      Top             =   600
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
      Height          =   2910
      ItemData        =   "Form5.frx":0004
      Left            =   1320
      List            =   "Form5.frx":0006
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
End
Attribute VB_Name = "Form5"
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
        a(i) = Int(Rnd * 999 + 1)
        List1.AddItem a(i)
    Next i
End Sub

Private Sub Command2_Click()
    Dim b#
    Cls
    For i = 1 To 50
        b = b + a(i)
    Next i
    Print "平均值为：" & Round(b / 50, 2)
End Sub

Private Sub Command3_Click()
    Dim i%, max%
    Cls
    max = a(1)
    For i = 1 To 50
        If max < a(i) Then max = a(i)
    Next i
    Print "最大值是：" & max
End Sub

Private Sub Command4_Click()
    Dim i%
    List2.Clear
    For i = 1 To 50
        If a(i) Mod 2 <> 0 Then List2.AddItem a(i)
    Next i
End Sub

Private Sub Command5_Click()
    Dim i%, j%
    List2.Clear
    For i = 1 To 50
        For j = 2 To a(i)
            If a(i) Mod j = 0 Then Exit For
        Next j
        If a(i) = j Then List2.AddItem a(i)
    Next i
End Sub

Private Sub Command6_Click()
    Unload Form5
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
