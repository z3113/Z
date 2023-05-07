VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "举例2：成绩处理"
   ClientHeight    =   5430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8055
   LinkTopic       =   "Form6"
   ScaleHeight     =   5430
   ScaleWidth      =   8055
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "返回"
      Height          =   495
      Left            =   6480
      TabIndex        =   3
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "统计结果"
      Height          =   495
      Left            =   6480
      TabIndex        =   2
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "显示成绩"
      Height          =   495
      Left            =   6480
      TabIndex        =   1
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "输入成绩"
      Height          =   495
      Left            =   6480
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim a(10) As Integer, i As Integer
Private Sub Command1_Click()
    For i = 1 To 10
        a(i) = Val(InputBox("请输入第" & i & "个同学的成绩", "成绩输入", 60))
    Next i
End Sub

Private Sub Command2_Click()
    Print "10个同学的成绩为(五个一行)："
    For i = 1 To 10
        Print a(i);
        If i Mod 5 = 0 Then Print
    Next i
End Sub

Private Sub Command3_Click()
    Dim b As Double, c As Integer
    For i = 1 To 10
        b = b + a(i)
    Next i
    b = b / 50
    For i = 1 To 10
        If a(i) > b Then c = c + 1
    Next i
    Print "平均分为" & b
    Print "高于平均分的人有" & c & "个"
End Sub

Private Sub Command4_Click()
    Unload Form6
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
