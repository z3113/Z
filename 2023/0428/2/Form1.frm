VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   5520
   ScaleWidth      =   8565
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "退出（&D）"
      Height          =   615
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "综合应用（&C）"
      Height          =   615
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "简单应用（&B）"
      Height          =   615
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "基本操作（&A）"
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3720
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Form1.Hide
    Form2.Show
End Sub

Private Sub Command2_Click()
    Form1.Hide
    Form3.Show
End Sub

Private Sub Command3_Click()
    Form1.Hide
    Form4.Show
End Sub

Private Sub Command4_Click()
    Unload Form1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("是否退出？", vbYesNo + 32, "退出") = vbNo Then Cancel = True
End Sub

Private Sub Label1_Click()

End Sub
