VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "字体字号"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7830
   LinkTopic       =   "Form3"
   ScaleHeight     =   7095
   ScaleWidth      =   7830
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "退出"
      Height          =   495
      Left            =   6240
      TabIndex        =   21
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "重置"
      Height          =   495
      Left            =   6240
      TabIndex        =   20
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   495
      Left            =   6240
      TabIndex        =   19
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   4680
      Width           =   5895
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      LargeChange     =   5
      Left            =   1200
      Max             =   60
      Min             =   10
      TabIndex        =   15
      Top             =   4080
      Value           =   10
      Width           =   4575
   End
   Begin VB.Frame Frame3 
      Caption         =   "字形"
      Height          =   1815
      Left            =   4440
      TabIndex        =   9
      Top             =   1800
      Width           =   1575
      Begin VB.CheckBox Check3 
         Caption         =   "下划线"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   1320
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         Caption         =   "倾斜"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   840
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "加粗"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "颜色"
      Height          =   1815
      Left            =   2280
      TabIndex        =   5
      Top             =   1800
      Width           =   1575
      Begin VB.OptionButton Option6 
         Caption         =   "红色"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton Option5 
         Caption         =   "蓝色"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   840
         Width           =   735
      End
      Begin VB.OptionButton Option4 
         Caption         =   "绿色"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   1320
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "字体"
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   1575
      Begin VB.OptionButton Option3 
         Caption         =   "仿宋"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   1320
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         Caption         =   "黑体"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   840
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "宋体"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "Form3.frx":0000
      Top             =   120
      Width           =   5895
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3240
      TabIndex        =   17
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "60"
      Height          =   255
      Left            =   5880
      TabIndex        =   16
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      Height          =   255
      Left            =   840
      TabIndex        =   14
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "字号："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   4080
      Width           =   735
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
    Text1.FontBold = Not Text1.FontBold
End Sub

Private Sub Check2_Click()
    Text1.FontItalic = Not Text1.FontItalic
End Sub

Private Sub Check3_Click()
    Text1.FontUnderline = Not Text1.FontUnderline
End Sub

Private Sub Command1_Click()
    Dim a$, b$, c$, d$
    a = Text1.Text
    If Option1.Value = True Then b = Option1.Caption
    If Option2.Value = True Then b = Option2.Caption
    If Option3.Value = True Then b = Option3.Caption
    If Option4.Value = True Then c = Option4.Caption
    If Option5.Value = True Then c = Option5.Caption
    If Option6.Value = True Then c = Option6.Caption
    If Check1.Value = 1 Then d = d & " " & Check1.Caption
    If Check2.Value = 1 Then d = d & " " & Check2.Caption
    If Check3.Value = 1 Then d = d & " " & Check3.Caption
    Text2.Text = "文字内容： " & a & vbCrLf & "字体: " & b & vbCrLf & "文字颜色: " & c & vbCrLf & "字形为:" & d
End Sub

Private Sub Command2_Click()
    Text1.Text = ""
    Text1.FontName = "宋体"
    Text1.ForeColor = vbBlack
    Option1.Value = False
    Option2.Value = False
    Option3.Value = False
    Option4.Value = False
    Option5.Value = False
    Option6.Value = False
    Check1.Value = 0
    Check2.Value = 0
    Check3.Value = 0
    HScroll1.Value = 10
End Sub

Private Sub Command3_Click()
    Unload Form3
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub

Private Sub HScroll1_Change()
    Text1.FontSize = HScroll1.Value
    Label4.Caption = HScroll1.Value
End Sub

Private Sub Option1_Click()
    Text1.FontName = "宋体"
End Sub

Private Sub Option2_Click()
    Text1.FontName = "黑体"
End Sub

Private Sub Option3_Click()
    Text1.FontName = "仿宋"
End Sub

Private Sub Option4_Click()
    Text1.ForeColor = vbGreen
End Sub

Private Sub Option5_Click()
    Text1.ForeColor = vbBlue
End Sub

Private Sub Option6_Click()
    Text1.ForeColor = vbRed
End Sub

Private Sub Text1_Change()
    Randomize
    Form3.BackColor = RGB(Int(Rnd * 256), Int(Rnd * 256), Int(Rnd * 256))
End Sub
