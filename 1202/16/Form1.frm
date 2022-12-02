VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9030
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":27A2
   ScaleHeight     =   4560
   ScaleWidth      =   9030
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command6 
      Cancel          =   -1  'True
      Caption         =   "关 闭 程 序 （ESC）"
      Height          =   495
      Left            =   5880
      TabIndex        =   6
      Top             =   3360
      Width           =   2535
   End
   Begin VB.CommandButton Command5 
      Caption         =   "开始做题序（ENTER）"
      Default         =   -1  'True
      Height          =   495
      Left            =   2880
      TabIndex        =   5
      Top             =   3360
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      Caption         =   "任务四（&D）"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6240
      TabIndex        =   4
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "任务三（&C）"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "任务二（&B）"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "任务一（&A）"
      Enabled         =   0   'False
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   24
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   8535
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
    Form1.Hide
    Form5.Show
End Sub

Private Sub Command5_Click()
    Randomize
    Form1.Picture = LoadPicture("")
    Form1.BackColor = RGB(Int(Rnd * 256), Int(Rnd * 256), Int(Rnd * 256))
    Command1.Enabled = True
    Command2.Enabled = True
    Command3.Enabled = True
    Command4.Enabled = True
End Sub

Private Sub Command6_Click()
    Unload Form1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MsgBox "今天我努力了，88！", 1 + 64, "提示"
End Sub
