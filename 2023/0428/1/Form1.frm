VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0E0FF&
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   8325
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "循环结构程序"
      Height          =   615
      Left            =   5160
      TabIndex        =   1
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "选择结构程序"
      Height          =   615
      Left            =   1320
      TabIndex        =   0
      Top             =   2280
      Width           =   1815
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

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("是否退出？", vbYesNo + 32, "退出") = vbNo Then Cancel = True
End Sub
