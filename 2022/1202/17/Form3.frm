VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "任务二"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "返回（ESC）"
      Height          =   495
      Left            =   3120
      TabIndex        =   9
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清除"
      Height          =   495
      Left            =   1680
      TabIndex        =   8
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "计算"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   3120
      TabIndex        =   6
      Text            =   "1"
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "半径（cm）"
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "周长（cm）"
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "面积（cm^2）"
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "r"
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   720
      Width           =   135
   End
   Begin VB.Line Line1 
      X1              =   960
      X2              =   1560
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Shape Shape1 
      Height          =   1215
      Left            =   360
      Shape           =   3  'Circle
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const pi! = 3.1415926
Option Explicit

Private Sub Command1_Click()
    Dim a!, b!, c!
    a = Text3.Text
    b = 2 * pi * a
    c = a * a * pi
    Text2.Text = b
    Text1.Text = c
End Sub

Private Sub Command2_Click()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
End Sub

Private Sub Command3_Click()
    Unload Form3
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
