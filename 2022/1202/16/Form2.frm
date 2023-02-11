VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "任务一：获取字符串"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6270
   LinkTopic       =   "Form2"
   ScaleHeight     =   7455
   ScaleWidth      =   6270
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   5400
      Width           =   5775
   End
   Begin VB.CommandButton Command7 
      Cancel          =   -1  'True
      Caption         =   "返回（ESC）"
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   6600
      Width           =   5775
   End
   Begin VB.CommandButton Command6 
      Caption         =   "从第5个开始取8个"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   4680
      Width           =   5775
   End
   Begin VB.CommandButton Command5 
      Caption         =   "去右边8个"
      Height          =   495
      Left            =   3480
      TabIndex        =   6
      Top             =   3840
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      Caption         =   "取左边5个"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   3840
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Text            =   "ABCDE123456FGHIJ67890"
      Top             =   2760
      Width           =   5775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "产生字符串""BBB    cccc     DDDDD""(中间分别4个和5个空格)"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   5775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "产生包含5个空格字符的字符串"
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   840
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "产生字符串“AAAA”"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5775
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      X1              =   120
      X2              =   6120
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      X1              =   120
      X2              =   6120
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   3
      X1              =   120
      X2              =   6120
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   3
      X1              =   120
      X2              =   6120
      Y1              =   2400
      Y2              =   2400
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Text1.Text = String(5, "A")
    MsgBox "产生的字符串是：" & Text1.Text & "！", 0 + 64, "获取字符串"
End Sub

Private Sub Command2_Click()
    Text1.Text = Space(5)
    MsgBox "产生的字符串是：" & Text1.Text & "！", 0 + 64, "获取字符串"
End Sub

Private Sub Command3_Click()
    Text1.Text = String(3, "B") & Space(4) & String(4, "c") & Space(5) & String(5, "D")
    MsgBox "产生的字符串是：" & Text1.Text & "！", 0 + 64, "获取字符串"
End Sub

Private Sub Command4_Click()
    Text3.Text = Left(Text2.Text, 5)
End Sub

Private Sub Command5_Click()
    Text3.Text = Right(Text2.Text, 8)
End Sub

Private Sub Command6_Click()
    Text3.Text = Mid(Text2.Text, 5, 8)
End Sub

Private Sub Command7_Click()
    Unload Form2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
