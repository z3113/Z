VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "主窗体"
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   ScaleHeight     =   4275
   ScaleWidth      =   9015
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command6 
      Cancel          =   -1  'True
      Caption         =   "关 闭 程 序 （ESC）"
      Height          =   495
      Left            =   600
      TabIndex        =   6
      Top             =   3120
      Width           =   7455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "任务五（&E）"
      Height          =   495
      Left            =   6840
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "任务四（&D）"
      Height          =   495
      Left            =   5280
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "任务三（&C）"
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "任务二（&B）"
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "任务一（&A）"
      Height          =   495
      Index           =   0
      Left            =   600
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "顺序结构上机练习"
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
      Height          =   735
      Left            =   2520
      TabIndex        =   0
      Top             =   480
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click(Index As Integer)
    Unload Form1
    Form2.Show
End Sub

Private Sub Command2_Click()
    Unload Form1
    Form3.Show
End Sub

Private Sub Command3_Click()
    Unload Form1
    Form4.Show
End Sub

Private Sub Command4_Click()
    Unload Form1
    Form5.Show
End Sub

Private Sub Command5_Click()
    Unload Form1
    Form6.Show
End Sub

Private Sub Command6_Click()
    End
End Sub
