VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "分支语句练习"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   4965
   ScaleWidth      =   8670
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command5 
      Caption         =   "退  出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   6960
      TabIndex        =   4
      Top             =   3600
      Width           =   1410
   End
   Begin VB.CommandButton Command4 
      Caption         =   "练习四"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   5280
      TabIndex        =   3
      Top             =   3600
      Width           =   1410
   End
   Begin VB.CommandButton Command3 
      Caption         =   "练习三"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   3600
      TabIndex        =   2
      Top             =   3600
      Width           =   1410
   End
   Begin VB.CommandButton Command2 
      Caption         =   "练习二"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1920
      TabIndex        =   1
      Top             =   3600
      Width           =   1410
   End
   Begin VB.CommandButton Command1 
      Caption         =   "练习一"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   240
      TabIndex        =   0
      Top             =   3600
      Width           =   1410
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
    End
End Sub
