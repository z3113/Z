VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "图形打印"
   ClientHeight    =   6135
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8535
   LinkTopic       =   "Form4"
   ScaleHeight     =   6135
   ScaleWidth      =   8535
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "图形3"
      Height          =   495
      Left            =   6960
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "图形2"
      Height          =   495
      Left            =   6960
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "图形1"
      Height          =   495
      Left            =   6960
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Cls
    Dim i%, j%, a%
    a = Val(InputBox("请输入不小于2的正整数", "输入", 5))
    If a >= 2 Then
        Print "1234567890123456789012345678901234567890"
        For i = 1 To a
            Print Tab(3 * (a - i) + 1);
            For j = 1 To i
                Print i;
            Next j
        Next i
    Else
        MsgBox "输入数字有误！", vbOKCancel + 16, "温馨提示"
    End If
End Sub

Private Sub Command2_Click()
    Cls
    Dim i%, j%, a%
    a = Val(InputBox("请输入不小于2的正整数", "输入", 5))
    If a >= 2 Then
        Print "1234567890123456789012345678901234567890"
        For i = 1 To a
            Print Tab(a - i + 1);
            For j = 1 To 2 * i - 1
                If j Mod 2 = 0 Then Print "*"; Else Print "$";
            Next j
        Next i
    Else
        MsgBox "输入数字错误！", vbOKCancel + 16, "温馨提示"
    End If
End Sub

Private Sub Command3_Click()
    Cls
    Dim i%, j%, a%
    a = Val(InputBox("请输入不小于2的正整数", "输入", 5))
    If a >= 2 Then
        Print "1234567890123456789012345678901234567890"
        For i = -a To a
            Print Tab(a - Abs(i) + 1);
            For j = -Abs(i) To Abs(i)
                Print "$";
            Next j
        Next i
    Else
        MsgBox "输入数字错误！", vbOKCancel + 16, "温馨提示"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
