VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6855
   LinkTopic       =   "Form3"
   ScaleHeight     =   4470
   ScaleWidth      =   6855
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "清除"
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "判合格（块IF）"
      Height          =   615
      Left            =   3480
      TabIndex        =   3
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "判合格(行IF)"
      Height          =   615
      Left            =   1080
      TabIndex        =   2
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "输入你期中考的成绩："
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim a!
    a = Val(Text1.Text)
    If a >= 60 Then Print "合格" Else Print "不合格"
End Sub

Private Sub Command2_Click()
    Dim a!
    a = Val(Text1.Text)
    If a >= 60 Then
        Print "合格！"
    Else
        Print "不合格！！！"
    End If
End Sub

Private Sub Command3_Click()
    Cls
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form2.Show
End Sub
