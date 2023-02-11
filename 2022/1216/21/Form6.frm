VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "征税"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7725
   LinkTopic       =   "Form6"
   ScaleHeight     =   4575
   ScaleWidth      =   7725
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "返回"
      Height          =   615
      Left            =   6000
      TabIndex        =   1
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   $"Form6.frx":0000
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
      Height          =   1815
      Left            =   360
      TabIndex        =   0
      Top             =   1200
      Width           =   7095
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Form6
End Sub

Private Sub Form_Activate()
    Dim a!
    a = Val(InputBox("请输入金额：", "税率计算器", 0))
    If a <= 800 Then
        Print "金额为：" & a & "应缴税款为：" & 0
    ElseIf a <= 1600 Then
        Print "金额为：" & a & "应缴税款为：" & (a - 800) * 0.05
    ElseIf a <= 3000 Then
        Print "金额为：" & a & "应缴税款为：" & 40 + (a - 1600) * 0.08
    Else
        Print "金额为：" & a & "应缴税款为：" & 152 + (a - 3000) * 0.1
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form2.Show
End Sub
