VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "for循环练习2"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12015
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   12015
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   6
      Top             =   3600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
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
      Left            =   10320
      TabIndex        =   5
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "正负累加"
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
      Left            =   8880
      TabIndex        =   4
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "n项之积"
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
      Left            =   7440
      TabIndex        =   3
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "n项之和"
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
      Left            =   8880
      TabIndex        =   2
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "随机数"
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
      Left            =   7440
      TabIndex        =   1
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "题目显示在这"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2895
      Left            =   5880
      TabIndex        =   0
      Top             =   240
      Width           =   5895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Randomize
    Dim i%
    Cls
    Label1.Caption = "1、单击随机数按钮，打印20个3位随机整数，每行5个。"
    For i = 1 To 20
        Print Int(Rnd * 900) + 100;
        If i Mod 5 = 0 Then Print
    Next i
End Sub

Private Sub Command2_Click()
    Dim i%, n%, a#
    Cls
    Label1.Caption = "2、使用inputbox输入项数n，输出以下数列前n项之和，四舍五入保留两位小数1+1/2+1/4+1/8+1/16+……"
    n = Int(InputBox("请输入项数n的值", "输入", 10))
    For i = 1 To n
        a = a + 1 / 2 ^ (i - 1)
    Next i
    a = Round(a, 2)
    Print "您输入的项数是："; n
    Print "1+1/2+1/4+1/8+1/16+……="; a
End Sub

Private Sub Command3_Click()
    Dim i%, n%, a#
    Cls
    Label1.Caption = "3、使用文本框输入项数n，输出以下数列前n项之积，四舍五入保四位小数（1/2）*（2/3）*（3/4）*（4/5） ……"
    Text1.Visible = True
    n = Val(Text1.Text)
    If n > 0 Then
        a = 1
        For i = 1 To n
            a = a * i / (i + 1)
        Next i
        a = Round(a, 4)
        Print "您输入的项数是："; n
        Print "(1/2)*(2/3)*(3/4)*(4/5)……="; a
    End If
End Sub

Private Sub Command4_Click()
    Dim i%, n%, a#, b%
    Cls
    Label1.Caption = "4、程序运行时要求在输入框中输入正整数N，若此输是[1，20]之间的，则求1-2/3+3/5-4/7+……+n/(2n-1),若不是，则显示错误信息"
    n = Int(InputBox("请输入项数n的值", "输入", 10))
    If 1 <= n <= 20 Then
        b = 1
        For i = 1 To n
            a = a + b * i / (2 * i + -1)
            b = -b
        Next i
        a = Round(a, 3)
        Print "您输入的项数是："; n
        Print "1-2/3+3/5-4/7+……+n/(2n-1)="; a
    Else
        MsgBox "n不在规定范围内！", vbOKCancel, "提示"
    End If
End Sub

Private Sub Command5_Click()
    Unload Form1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("真的要退出吗？", vbYesNo, "提示") = vbNo Then Cancel = True
End Sub
