VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "for循环练习5"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7815
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   7815
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command7 
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   7
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "任务六"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   6
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "任务五"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   5
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "任务四"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   4
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "任务三"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   3
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "任务二"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   2
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "任务一"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   3000
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Cls
    Dim i%, a&
    Text1.Text = "1、将1+3^2+3^3+3^4+3^5+……+3^10 之和用消息框输出。"
    a = 1
    For i = 2 To 10
        a = a + 3 ^ i
    Next i
    Print "1+3^2+3^3+3^4+3^5+……+3^10="; a
    MsgBox "1+3^2+3^3+3^4+3^5+……+3^10=" & a, vbOKCancel + 64, "求和结果"
End Sub

Private Sub Command2_Click()
    Cls
    Dim i%, a%, b#
    b = 1
    Text1.Text = "2、inputbox输入项数n，输出以下数列的前n项之积。"
    a = Val(InputBox("请输入n的值", "输入", 10))
    For i = 1 To a
        b = b * 2 ^ (i - 1) / (i + 1)
    Next i
    Print "n="; a; "积="; b
End Sub

Private Sub Command3_Click()
    Cls
    Dim i%, n%, a%, b%, c%, d&
    a = 1
    b = 1
    d = 2
    Text1.Text = "3、有一数列，a1,a2,a3…an,其中a1=1,a2=1,ai满足条件ai=ai-1+ai-2，从键盘输入项数，求第n项的值和前n项的和。"
    n = Val(InputBox("请输入n的值", "输入", 10))
    If n >= 3 Then
        Print a; b;
        For i = 3 To n
            c = a + b
            d = d + c
            Print c;
            If i Mod 3 = 0 Then Print
            a = b
            b = c
        Next i
    Else
        MsgBox "请输入不小于3的正整数", vbOKCancel + 48, "温馨提示："
    End If
    Print "第"; n; "项的值为："; c
    Print "前"; n; "项的和为："; d
End Sub

Private Sub Command4_Click()
    Cls
    Dim i%, a%, b%, min%, max%, e&
    min = 100
    max = 0
    Text1.Text = "4、 判定全班同学的成绩等级（班级人数用输入框输入）"
    a = Val(InputBox("请输入班级人数", "输入人数", 10))
    For i = 1 To a
        b = InputBox("请输入第" & i & "个同学的成绩", "输入成绩", 60)
        If b < 60 Then
            Print "第"; i; "个同学的分数是："; b
            Print "成绩等级为：不及格"
        ElseIf b < 80 Then
            Print "第"; i; "个同学的分数是："; b
            Print "成绩等级为：及格"
        ElseIf b < 90 Then
            Print "第"; i; "个同学的分数是："; b
            Print "成绩等级为：良好"
        Else
            Print "第"; i; "个同学的分数是："; b
            Print "成绩等级为：优秀"
        End If
        e = e + b
        If b >= max Then max = b
        If b <= min Then min = b
    Next i
    Print "班级最高分为："; max
    Print "班级最低分为："; min
    Print "班级平均分为："; e
End Sub

Private Sub Command5_Click()
    Cls
    Dim i%, a%, b$, c$, d$
    d = ""
    Text1.Text = "5、任意输入N个字符，倒打印原字符。"
    b = InputBox("请输入一个字符串", "输入", "abcdef")
    a = Len(b)
    For i = a To 1 Step -1
        c = Mid(b, i, 1)
        d = d & c
    Next i
    Print "原字符串为："; b
    Print "倒置后的字符串为："; d
End Sub

Private Sub Command6_Click()
    Cls
    Dim i%, a!, b!
    a = 200
    b = -200
    Text1.Text = "5、一小球从 200 米高度自由下落，每次落地后反弹原高度的一半，然后再落下……，求该小球第十次落地时共经过了多少米的路程？"
    For i = 1 To 10
        b = b + 2 * a
        a = a / 2
    Next i
    Print "第10次落地经过了"; b; "米"
End Sub

Private Sub Command7_Click()
    Unload Form1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("确定退出吗？", vbOKCancel + 64, "退出提示") = vbCancel Then Cancel = True
End Sub
