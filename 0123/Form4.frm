VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "综合应用"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5565
   LinkTopic       =   "Form4"
   ScaleHeight     =   3285
   ScaleWidth      =   5565
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text4 
      Height          =   615
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   705
      Left            =   1920
      TabIndex        =   9
      Top             =   2280
      Width           =   3300
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   1920
      TabIndex        =   8
      Top             =   240
      Width           =   3300
   End
   Begin VB.Label Label5 
      Caption         =   "重新组合的三位数(y):"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "个位(g):"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   5
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "十位(s):"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      TabIndex        =   3
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "百位(b):"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "请输入三位数（x）："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x%, y%, b&, s%, g%

Private Sub Form_Activate()
    x = Val(InputBox("请输入三位整数", "键盘输入数", 123))
    b = x \ 100
    s = x \ 10 Mod 10
    g = x Mod 10
    y = b * 100 + s * 10 + g
    If x < 100 Or x > 999 Then MsgBox "你输入的数不是三位的，请重输！"
    Label6.Caption = x
    Text2.Text = b
    Text3.Text = s
    Text4.Text = g
    Label7.Caption = y
End Sub

Private Sub Form_DblClick()
    Unload Form4
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub

