VERSION 5.00
Begin VB.Form Form4 
   ClientHeight    =   7395
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6120
   LinkTopic       =   "Form4"
   ScaleHeight     =   7395
   ScaleWidth      =   6120
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "周长"
      Height          =   495
      Left            =   4680
      TabIndex        =   16
      Top             =   4680
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3015
      Left            =   120
      TabIndex        =   10
      Top             =   3720
      Width           =   5655
      Begin VB.CommandButton Command4 
         Caption         =   "面积"
         Height          =   495
         Left            =   4560
         TabIndex        =   17
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         Height          =   615
         Left            =   2040
         TabIndex        =   15
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   2040
         TabIndex        =   13
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "结果："
         Height          =   615
         Left            =   360
         TabIndex        =   14
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "输入圆的半径："
         Height          =   495
         Left            =   360
         TabIndex        =   12
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "二、请在文本框1输入圆的半径，计算圆的周长和面积显示在第二个文本框中。（第二个文本框不可用）"
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   5415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin VB.CommandButton Command2 
         Caption         =   "面积"
         Height          =   615
         Left            =   4560
         TabIndex        =   9
         Top             =   2160
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "周长"
         Height          =   615
         Left            =   4560
         TabIndex        =   8
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Height          =   615
         Left            =   3600
         TabIndex        =   5
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Height          =   615
         Left            =   2400
         TabIndex        =   4
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   615
         Left            =   1320
         TabIndex        =   2
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "结果："
         Height          =   615
         Left            =   360
         TabIndex        =   7
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   1320
         TabIndex        =   6
         Top             =   2040
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "三边长："
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "一、在文本框分别输入三边长，计算三角形的周长和面积。海伦公式：p=1/2(a+b+c) s=p(p-a)(p-b)(p-c)的平方根。 "
         Height          =   495
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   5055
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Const Pi = 3.1415926
 Option Explicit



Private Sub Command1_Click()
    Dim a!, b!, c!
    a = Val(Text1.Text)
    b = Val(Text2.Text)
    c = Val(Text3.Text)
    If a + b > c And a + c > b And b + c > a Then
        Label3.Caption = a + b + c
    Else
        MsgBox "", 1 + 16, ""
    End If
End Sub

Private Sub Command2_Click()
    Dim a!, b!, c!, p!
    a = Val(Text1.Text)
    b = Val(Text2.Text)
    c = Val(Text3.Text)
    If a + b > c And a + c > b And b + c > a Then
        p = (a + b + c) / 2
        Label3.Caption = Sqr(p * (p - a) * (p - b) * (p - c))
    Else
        MsgBox "", 1 + 16, ""
    End If
End Sub

Private Sub Command3_Click()
    Dim zc!
    zc = 2 * Pi * Text4.Text
    Text5.Text = zc
End Sub

Private Sub Command4_Click()
    Dim mj!
    mj = Pi * Text4.Text * Text4.Text
    Text5.Text = mj
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Form4
    Form1.Show
End Sub
