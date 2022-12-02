VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "任务三"
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4140
   LinkTopic       =   "Form4"
   ScaleHeight     =   4275
   ScaleWidth      =   4140
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "返回（ESC）"
      Height          =   495
      Left            =   2160
      TabIndex        =   9
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "计算"
      Height          =   495
      Left            =   600
      TabIndex        =   8
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   495
      Left            =   2280
      TabIndex        =   7
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Text            =   "0"
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "平方根（C）："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "立方值（B）："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "平方值（A）："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "输入一个正数"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim a!, b!, c!, d!
    d = Val(Text1.Text)
    a = d * d
    b = d * d * d
    c = Sqr(d)
    Text2.Text = a
    Text3.Text = b
    Text4.Text = c
End Sub

Private Sub Command2_Click()
    Unload Form4
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub

