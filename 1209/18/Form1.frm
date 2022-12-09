VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "开车比赛"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8160
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   8160
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5400
      Top             =   960
   End
   Begin VB.CommandButton Command3 
      Caption         =   "停止"
      Height          =   495
      Left            =   4440
      TabIndex        =   4
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "自动挡开车"
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "手动挡开车"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   960
      Picture         =   "Form1.frx":0B3A
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   1
      Top             =   2400
      Width           =   600
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   960
      Picture         =   "Form1.frx":0F7D
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   0
      Top             =   720
      Width           =   600
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Picture1.Left = Picture1.Left + 100
End Sub

Private Sub Command2_Click()
    Timer1.Enabled = True
End Sub

Private Sub Command3_Click()
    Timer1.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form2.Show
End Sub

Private Sub Timer1_Timer()
    Picture2.Left = Picture2.Left + 100
    If Picture2.Left >= Form1.ScaleWidth Then
        Picture2.Left = -Picture2.Width
    End If
End Sub
