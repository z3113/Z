VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "循环结构"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6165
   LinkTopic       =   "Form3"
   ScaleHeight     =   5430
   ScaleWidth      =   6165
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   3000
      TabIndex        =   6
      Top             =   2880
      Width           =   3015
   End
   Begin VB.CommandButton Command6 
      Caption         =   "第六题"
      Height          =   495
      Left            =   4560
      TabIndex        =   5
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "第五题"
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "第四题"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "第三题"
      Height          =   495
      Left            =   4560
      TabIndex        =   2
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "第二题"
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "第一题"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "第三题输入"
      Height          =   180
      Left            =   1920
      TabIndex        =   7
      Top             =   3120
      Width           =   900
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Cls
    Dim i%, a%, b!
    a = 1
    For i = 1 To 200
        b = b + a / i
        a = -a
    Next i
    Print "前200项之和为" & b
End Sub

Private Sub Command2_Click()
    Cls
    Dim i%, a$, b$
    a = InputBox("请输入一个字符串")
    For i = Len(a) To 1 Step -1
        b = b & Mid(a, i, 1)
    Next i
    If b = a Then
        Print a & "是回文"
    Else
        Print a & "不是回文"
    End If
End Sub

Private Sub Command3_Click()
    Cls
    Dim i%, a$, b%, c%, d%, e%
    For i = 1 To Len(Text1.Text)
        a = Mid(Text1.Text, i, 1)
        If "A" <= a And a <= "Z" Then
            b = b + 1
        ElseIf "A" <= a And a <= "z" Then
            c = c + 1
        ElseIf "0" <= a And a <= "9" Then
            d = d + 1
        Else
            e = e + 1
        End If
    Next i
    Print "大写英文个数" & b
    Print "小写英文个数" & c
    Print "数字个数" & d
    Print "其他字符个数" & e
End Sub

Private Sub Command4_Click()
    Cls
    Dim i%, a$, b%
    For i = 1 To 10
        a = InputBox("请输入第" & i & "个字符串")
        If Mid(a, 1, 1) = "D" Then b = b + 1
    Next i
    Print "以字母D开头的单词有" & b & "个"
End Sub

Private Sub Command5_Click()
    Cls
    Dim i%
    For i = 100 To 999
        If (i \ 100) ^ 3 + (i \ 10 Mod 10) ^ 3 + (i Mod 10) ^ 3 = i Then
            Print i;
        End If
    Next i
End Sub

Private Sub Command6_Click()
    Dim i%, j%, k%, a%, b%
    For i = 1 To 9
        For j = 0 To 9
            For k = 0 To 9
                a = i & j & k
                If (i = 2 Or j = 2 Or k = 2) And a Mod 9 = 0 Then
                    Print a,
                    b = b + 1
                    If b Mod 7 = 0 Then Print
                End If
            Next k
        Next j
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
