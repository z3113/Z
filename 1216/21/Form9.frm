VERSION 5.00
Begin VB.Form Form9 
   Caption         =   "提高"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5475
   LinkTopic       =   "Form9"
   ScaleHeight     =   3060
   ScaleWidth      =   5475
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "返回"
      Height          =   495
      Left            =   3480
      TabIndex        =   1
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "从键盘输出一个不大于5位的正整数，求出它是几位数，将该数位逆序输出。"
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
      Height          =   855
      Left            =   1320
      TabIndex        =   0
      Top             =   1320
      Width           =   3735
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Form9
End Sub

Private Sub Form_Activate()
    Dim a%
    a = Val(InputBox("请输入一个不大于5位的正整数：", "判断数位"))
    If a >= 10000 Then
        Print "此数是5位数"
        Print StrReverse(a)
    ElseIf a >= 1000 Then
        Print "此数是4位数"
        Print StrReverse(a)
    ElseIf a >= 100 Then
        Print "此数是3位数"
        Print StrReverse(a)
    ElseIf a >= 10 Then
        Print "此数位2位数"
        Print StrReverse(a)
    ElseIf a >= 0 Then
        Print "此数为1位数"
        Print StrReverse(a)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form2.Show
End Sub
