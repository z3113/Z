VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "成绩统计"
   ClientHeight    =   4230
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8790
   LinkTopic       =   "Form5"
   ScaleHeight     =   4230
   ScaleWidth      =   8790
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "统计"
      Height          =   495
      Left            =   7440
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "输入成绩"
      Height          =   495
      Left            =   7440
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "5、单击按钮Command1后，用户利用输入框连续输入10个学生的成绩，单击按钮Command2后统计所有学生的平均分，并显示最低分及这些人的位置。"
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
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   8535
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim a(10) As Integer, i As Integer

Private Sub Command1_Click()
    Cls
    Print "十个同学的成绩如下："
    For i = 1 To 10
        a(i) = Val(InputBox("请输入第" & i & "个同学的成绩", "输入成绩"))
        Print a(i);
    Next i
End Sub

Private Sub Command2_Click()
    Dim b%, min%
    min = a(1)
    For i = 1 To 10
        b = b + a(i)
        If min >= a(i) Then min = a(i)
    Next i
    Print "平均分为："; Round(b / 10, 1); "分"
    Print "最低分是："; min; "分，他是第";
    For i = 1 To 10
        If a(i) = min Then Print i;
    Next i
    Print "个同学"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
