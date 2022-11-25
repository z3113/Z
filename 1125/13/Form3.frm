VERSION 5.00
Begin VB.Form Form3 
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6060
   LinkTopic       =   "Form3"
   ScaleHeight     =   5415
   ScaleWidth      =   6060
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command6 
      Height          =   495
      Left            =   4560
      TabIndex        =   15
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Height          =   495
      Left            =   4560
      TabIndex        =   14
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1200
      TabIndex        =   9
      Top             =   3240
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Height          =   495
      Left            =   3240
      TabIndex        =   8
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Height          =   495
      Left            =   3240
      TabIndex        =   7
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "一、字符串练习：请在输入框里输入一个多于5个字符的文本"
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   5055
   End
   Begin VB.Label Label8 
      Caption         =   "输入："
      Height          =   495
      Left            =   240
      TabIndex        =   12
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "输出："
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   1200
      TabIndex        =   10
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "二、数值练习：请在输入框里输入一个实数"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   2520
      Width           =   5055
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "输出："
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "输入："
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   615
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim a1$
    a1 = Text1.Text
    Label4.Caption = a1
End Sub

Private Sub Command2_Click()
    Dim a2 As String * 5
    a2 = Text1.Text
    Label4.Caption = a2
End Sub

Private Sub Command3_Click()
    Dim a%
    a = Val(Text2.Text)
    Label6.Caption = a
End Sub

Private Sub Command4_Click()
    Dim b&
    b = Val(Text2.Text)
    Label6.Caption = b
End Sub

Private Sub Command5_Click()
    Dim c!
    c = Val(Text2.Text)
    Label6.Caption = c
End Sub

Private Sub Command6_Click()
    Dim d#
    d = Val(Text2.Text)
    Label6.Caption = d
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Form3
    Form1.Show
End Sub
