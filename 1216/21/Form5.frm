VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "商场打折"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6885
   LinkTopic       =   "Form5"
   ScaleHeight     =   3855
   ScaleWidth      =   6885
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "返回"
      Height          =   495
      Left            =   5400
      TabIndex        =   1
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form5.frx":0000
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
      Height          =   3255
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   6495
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Form5
End Sub

Private Sub Form_Activate()
    Dim a!
    a = Val(InputBox("请输入购货款"))
    If a < 250 Then
        Print "付款为：" & a
    ElseIf a < 500 Then
        Print "付款为：" & a * 0.95
    ElseIf a < 1000 Then
        Print "付款为：" & a * 0.925
    ElseIf a < 2000 Then
        Print "付款为：" & a * 0.9
    Else
        Print "付款为：" & a * 0.85
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form2.Show
End Sub
