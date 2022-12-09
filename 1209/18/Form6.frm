VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "计算"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8655
   LinkTopic       =   "Form6"
   ScaleHeight     =   4875
   ScaleWidth      =   8655
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "从小到大输出两个整数"
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   4200
      Width           =   7695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "计算苹果价格！（行IF）"
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   2400
      Width           =   6135
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Text            =   "2"
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "弹出对话框输入成绩（整数），判定后用对话输出是否合格？（块IF）"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   8175
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   8520
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   8520
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label3 
      Caption         =   "输入两个整数，分别放在x和y变量中，比较它们大小，然后将大数放在x中，小数放在y中。"
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
      Left            =   240
      TabIndex        =   3
      Top             =   3480
      Width           =   8055
   End
   Begin VB.Label Label2 
      Caption         =   $"Form6.frx":0000
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
      Height          =   1095
      Left            =   2040
      TabIndex        =   2
      Top             =   960
      Width           =   6255
   End
   Begin VB.Label Label1 
      Caption         =   "苹果重量："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim a%
    a = Val(InputBox("请输入成绩(整数)", "输入成绩", 60))
    If a >= 60 Then
        MsgBox "恭喜你合格了！"
    Else
        MsgBox "很遗憾，你没有合格。"
    End If
End Sub

Private Sub Command2_Click()
    Dim a!
    a = Val(Text1.Text)
    If a < 2 Then MsgBox "你买了" & a & "千克苹果，共计" & a * 1.5 & "元" Else MsgBox "你买了" & a & "千克苹果，共计" & a * 1.5 * 0.8 & "元"
End Sub

Private Sub Command3_Click()
    Dim x%, y%, z%
    x = InputBox("请为x赋值", "数据输入", 100)
    y = InputBox("请为y赋值", "数据输入", 200)
    If x < y Then
        z = x
        x = y
        y = z
    End If
    MsgBox "两个数为" & y & "和" & x
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form2.Show
End Sub
