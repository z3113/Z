VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8295
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":08CA
   ScaleHeight     =   6015
   ScaleWidth      =   8295
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command5 
      Caption         =   "Unload"
      Height          =   495
      Left            =   6360
      TabIndex        =   4
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Discount"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4320
      TabIndex        =   3
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Timer and scroll"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Font setting"
      Enabled         =   0   'False
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "START"
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   1800
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Command1.Visible = False
    Command2.Enabled = True
    Command3.Enabled = True
    Command4.Enabled = True
End Sub

Private Sub Command2_Click()
    Form2.Show
End Sub

Private Sub Command3_Click()
    Form3.Show
End Sub

Private Sub Command4_Click()
    Form4.Show
End Sub

Private Sub Command5_Click()
    Unload Form1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MsgBox("确定要关闭窗口吗？", vbOKCancel + 32, "关闭") = vbCancel Then Cancel = True
End Sub
