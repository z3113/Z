VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "分段函数"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7530
   LinkTopic       =   "Form8"
   ScaleHeight     =   6135
   ScaleWidth      =   7530
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "计算"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   4
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Y="
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   2
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "X="
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   0
      Top             =   840
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   5940
      Left            =   240
      Picture         =   "Form8.frx":0000
      Top             =   240
      Width           =   4155
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim x!
    x = Val(Text1.Text)
    If x < 0 Then
        Text2.Text = 0
    ElseIf x < 1 Then
        Text2.Text = 1
    ElseIf x < 2 Then
        Text2.Text = 2
    Else
        Text2.Text = 3
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form2.Show
End Sub
