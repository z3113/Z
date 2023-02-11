VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "数位处理"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7545
   LinkTopic       =   "Form7"
   ScaleHeight     =   3765
   ScaleWidth      =   7545
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "返回"
      Height          =   495
      Left            =   4560
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "从键盘输入一个不大于5位的正整数，求出它是几位数，并把每一位数输出。"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   1800
      TabIndex        =   0
      Top             =   2040
      Width           =   5295
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Form7
End Sub

Private Sub Form_Activate()
    Dim a%
    a = Val(InputBox("请输入一个不大于5位的正整数：", "判断数位"))
    If a >= 10000 Then
        Print "此数是5位数"
        Print a \ 10000; a \ 1000 Mod 10; a \ 100 Mod 10; a \ 10 Mod 10; a Mod 10
    ElseIf a >= 1000 Then
        Print "此数是4位数"
        Print a \ 1000; a \ 100 Mod 10; a \ 10 Mod 10; a Mod 10
    ElseIf a >= 100 Then
        Print "此数是3位数"
        Print a \ 100; a \ 10 Mod 10; a Mod 10
    ElseIf a >= 10 Then
        Print "此数位2位数"
        Print a \ 10; a Mod 10
    ElseIf a >= 0 Then
        Print "此数为1位数"
        Print a
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form2.Show
End Sub
