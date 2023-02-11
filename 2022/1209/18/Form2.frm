VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "IF语句综合1"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8760
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0B3A
   ScaleHeight     =   5070
   ScaleWidth      =   8760
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command7 
      Caption         =   "退    出"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   7
      Top             =   3480
      Width           =   2775
   End
   Begin VB.CommandButton Command6 
      Caption         =   "练习六"
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
      Left            =   360
      TabIndex        =   6
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "练习五"
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
      Left            =   7080
      TabIndex        =   5
      Top             =   1560
      Width           =   1335
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
      Height          =   495
      Left            =   5400
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
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
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
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
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
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
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "IF语句学习"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   3000
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Form2
    Form3.Show
End Sub

Private Sub Command2_Click()
    Unload Form2
    Form4.Show
End Sub

Private Sub Command3_Click()
    Unload Form2
    Form5.Show
End Sub

Private Sub Command4_Click()
    Unload Form2
    Form6.Show
End Sub

Private Sub Command5_Click()
    Unload Form2
    Form7.Show
End Sub

Private Sub Command6_Click()
    Unload Form2
    Form8.Show
End Sub

Private Sub Command7_Click()
    End
End Sub
