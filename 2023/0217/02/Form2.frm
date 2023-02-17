VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00C0E0FF&
   Caption         =   "任务六：求各数位之和"
   ClientHeight    =   3045
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5550
   LinkTopic       =   "Form2"
   ScaleHeight     =   3045
   ScaleWidth      =   5550
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "返回"
      Height          =   495
      Left            =   4200
      TabIndex        =   7
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清空"
      Height          =   495
      Left            =   4200
      TabIndex        =   6
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "计算"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4200
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2040
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   960
      MaxLength       =   9
      TabIndex        =   2
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "输出:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "输入:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "6、输入一个正整数，求各位数字之和"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim i%, j%, a&, b%
    a = Val(Text1.Text)
    For i = 1 To 9
        If a \ (10 ^ i) = 0 Then Exit For
    Next i
    For j = 1 To i
        b = b + a Mod 10
        a = a \ 10
    Next j
    Text2.Text = b
End Sub

Private Sub Command2_Click()
    Text1.Text = ""
    Text2.Text = ""
End Sub

Private Sub Command3_Click()
    Unload Form2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub

Private Sub Text1_Change()
    Command1.Enabled = True
End Sub
