VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "分段函数"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8535
   LinkTopic       =   "Form4"
   ScaleHeight     =   4350
   ScaleWidth      =   8535
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Height          =   975
      Left            =   5040
      Picture         =   "Form4.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   360
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   1680
      TabIndex        =   3
      Text            =   "结果"
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   1680
      TabIndex        =   1
      Text            =   "-2"
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Y="
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "X="
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim a!
    a = Val(Text1.Text)
    If a <> -2 Then
        Text2.Text = Sqr(a * a + 1)
    Else
        Text2.Text = a * a * a + 3 * a
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form2.Show
End Sub
