VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00C0E0FF&
   Caption         =   "综合应用"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7470
   LinkTopic       =   "Form4"
   ScaleHeight     =   4470
   ScaleWidth      =   7470
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "清空"
      Height          =   495
      Left            =   5520
      TabIndex        =   11
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "大小字符"
      Height          =   495
      Left            =   5520
      TabIndex        =   10
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "反向字符串"
      Height          =   495
      Left            =   5520
      TabIndex        =   9
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "随即生成"
      Height          =   495
      Left            =   5520
      TabIndex        =   8
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   7
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   5
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2040
      Width           =   5175
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "大小之差："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   6
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "最下字符："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "最大字符："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   1215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Long
Dim max
Dim min
Dim sj(200)

Private Sub Command1_Click()
    i = InputBox("请输入要生成字符串的长度（1-200）", "随机生成", 10)
    Randomize
    min = 200
    max = 10
    For i = 1 To i
        sj(i) = Int(Rnd * 191 + 10)
        Text1.Text = Text1.Text & Chr(sj(i))
        If max < sj(i) Then max = sj(i)
        If min > sj(i) Then min = sj(i)
    Next
End Sub

Private Sub Command2_Click()
    Text2.Text = StrReverse(Text1.Text)
End Sub

Private Sub Command3_Click()
    Text3.Text = Chr(max)
    Text4.Text = Chr(min)
    Text5.Text = max - min
End Sub

Private Sub Command4_Click()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
