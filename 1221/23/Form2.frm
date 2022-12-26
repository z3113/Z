VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "练习一、三分支函数"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8070
   LinkTopic       =   "Form2"
   ScaleHeight     =   4335
   ScaleWidth      =   8070
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "返回主窗体"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   3360
      Width           =   2895
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Select Case解决"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   3360
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ELSEIF解决"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   1
      Top             =   2520
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "普通IF解决"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   2520
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   480
      Picture         =   "Form2.frx":0000
      Top             =   240
      Width           =   4155
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim x!, y%
    x = Val(InputBox("请输入一个整数x：", "数据输入", 100))
    If x > 0 Then y = 1
    If x = 0 Then y = 0
    If x < 0 Then y = -1
    Print "y ="; y
End Sub

Private Sub Command2_Click()
    Dim x!, y%
    x = Val(InputBox("请输入一个整数x：", "数据输入", 0))
    If x > 0 Then
        y = 1
    ElseIf x = 0 Then
        y = 0
    ElseIf x < 0 Then
        y = -1
    End If
    Print "y ="; y
End Sub

Private Sub Command3_Click()
    Dim x!, y%
    x = Val(InputBox("请输入一个整数x：", "数据输入", -100))
    Select Case x
        Case Is > 0
            y = 1
        Case Is = 0
            y = 0
        Case Is < 0
            y = -1
    End Select
    Print "y ="; y
End Sub

Private Sub Command4_Click()
    Unload Form2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
