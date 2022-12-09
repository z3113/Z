VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "三角形计算"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   17655
   LinkTopic       =   "Form7"
   ScaleHeight     =   6090
   ScaleWidth      =   17655
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      BackColor       =   &H000000FF&
      Caption         =   "返回"
      Height          =   495
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "清除"
      Height          =   495
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "计算"
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   10
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   8
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   6
      Text            =   "5"
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   4
      Text            =   "3"
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Text            =   "2"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   8280
      Picture         =   "Form7.frx":0000
      ScaleHeight     =   4935
      ScaleWidth      =   9015
      TabIndex        =   0
      Top             =   360
      Width           =   9015
   End
   Begin VB.Label Label5 
      Caption         =   "周长：（保留2位小数）"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3960
      TabIndex        =   9
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "面积：（保留2位小数）"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "c"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   4800
      TabIndex        =   5
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "b"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   360
      Width           =   495
   End
   Begin VB.Line Line3 
      X1              =   3240
      X2              =   6360
      Y1              =   3000
      Y2              =   1080
   End
   Begin VB.Line Line2 
      X1              =   3240
      X2              =   1560
      Y1              =   3000
      Y2              =   1080
   End
   Begin VB.Line Line1 
      X1              =   1560
      X2              =   6360
      Y1              =   1080
      Y2              =   1080
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim a!, b!, c!, p!
    a = Val(Text1.Text)
    b = Val(Text2.Text)
    c = Val(Text3.Text)
    If a + b > c And a + c > b And b + c > a Then
        p = (a + b + c) / 2
        Text4.Text = Sqr(p * (p - a) * (p - b) * (p - c))
        Text5.Text = a + b + c
    Else
        MsgBox "无法构成三角形！", 0 + 48, "数据错误"
    End If
End Sub

Private Sub Command2_Click()
    Text4.Text = ""
    Text5.Text = ""
End Sub

Private Sub Command3_Click()
    Unload Form7
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form2.Show
End Sub
