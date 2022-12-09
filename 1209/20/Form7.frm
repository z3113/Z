VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "IF综合运用"
   ClientHeight    =   4140
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6645
   LinkTopic       =   "Form7"
   ScaleHeight     =   4140
   ScaleWidth      =   6645
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "4、输入三个数，从大到小排序输出"
      Height          =   735
      Left            =   3360
      TabIndex        =   3
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "3、输入一个任意数（含小数）判断奇偶性"
      Height          =   735
      Left            =   600
      TabIndex        =   2
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2、输入重量求运费  （块IF）"
      Height          =   735
      Left            =   3360
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1、判断胖瘦改版"
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim a!, b!
    a = Val(InputBox("输入身高（cm）", "输入", 170))
    b = Val(InputBox("输入体重（kg）", "输入", 80))
    If b / (a * a) < 30 Then
        MsgBox "你不偏胖"
    Else
        MsgBox "你偏胖了！"
    End If
End Sub

Private Sub Command2_Click()
    Dim a!
    a = Val(InputBox("请输入重量", "输入", 60))
    If a <= 50 Then
        MsgBox "运费是" & a * 0.13
    Else
        MsgBox "运费是" & 50 * 0.13 + (a - 50) * 0.2
    End If
End Sub

Private Sub Command3_Click()
    Dim a%
    a = InputBox("请输入一个整数", "判断奇偶输入", 100)
    If a Mod 2 = 0 Then
        MsgBox "是偶数"
    Else
        MsgBox "是奇数"
    End If
End Sub

Private Sub Command4_Click()
    Dim a!, b!, c!, t!
    a = Val(InputBox("first number", "enter", 30))
    b = Val(InputBox("second number", "enter", 10.5))
    c = Val(InputBox("third number", "enter", 50))
    If a > b Then
        t = a
        a = b
        b = t
    End If
    If a > c Then
        t = a
        a = c
        c = t
    End If
    If b > c Then
        t = b
        b = c
        c = t
    End If
    Print "三个数由小到大为" & a; b; c
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
