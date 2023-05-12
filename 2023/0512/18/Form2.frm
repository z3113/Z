VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "斐波那契数列"
   ClientHeight    =   4485
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7665
   LinkTopic       =   "Form2"
   ScaleHeight     =   4485
   ScaleWidth      =   7665
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "显示方式二"
      Height          =   495
      Left            =   5880
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "显示方式一"
      Height          =   495
      Left            =   5880
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "1、求斐波那契数列前20项的值的总和（用数组的方法实现）"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   840
      TabIndex        =   2
      Top             =   4080
      Width           =   5985
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim a(1 To 20) As Long, b As Long, i As Integer
    Cls
    a(1) = 1
    a(2) = 1
    b = a(1) + a(2)
    Print a(1); a(2);
    For i = 3 To 20
        a(i) = a(i - 2) + a(i - 1)
        Print a(i);
        If i Mod 5 = 0 Then Print
        b = b + a(i)
    Next i
    Print
    Print "总和为："; b
End Sub

Private Sub Command2_Click()
    Dim a(1 To 20) As Long, b As Long, i As Integer, j As Integer
    Cls
    a(1) = 1
    a(2) = 1
    b = 2
    For i = 1 To 2
        For j = 1 To i
            Print a(j);
        Next j
        Print
    Next i
    For i = 3 To 20
        a(i) = a(i - 2) + a(i - 1)
        For j = 1 To i
            Print a(j);
        Next j
        Print
        b = b + a(i)
    Next i
    Print
    Print "总和为："; b
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
