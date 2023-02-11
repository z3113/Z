VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "任务五"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7920
   LinkTopic       =   "Form6"
   ScaleHeight     =   4695
   ScaleWidth      =   7920
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "返回（ESC）"
      Height          =   495
      Left            =   4440
      TabIndex        =   12
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清除"
      Height          =   495
      Left            =   2640
      TabIndex        =   11
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "计算"
      Height          =   495
      Left            =   840
      TabIndex        =   10
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   4800
      TabIndex        =   9
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   2160
      TabIndex        =   7
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   5040
      TabIndex        =   5
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "周长："
      Height          =   255
      Left            =   3840
      TabIndex        =   8
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "面积："
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "c"
      Height          =   255
      Left            =   4560
      TabIndex        =   4
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "b"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "a"
      Height          =   255
      Left            =   2400
      TabIndex        =   0
      Top             =   240
      Width           =   255
   End
   Begin VB.Line Line3 
      X1              =   3240
      X2              =   6240
      Y1              =   2520
      Y2              =   840
   End
   Begin VB.Line Line2 
      X1              =   1320
      X2              =   3240
      Y1              =   840
      Y2              =   2520
   End
   Begin VB.Line Line1 
      X1              =   1320
      X2              =   6240
      Y1              =   840
      Y2              =   840
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim a!, b!, c!, p!
    a = Text1.Text
    b = Text2.Text
    c = Text3.Text
    If a + b > c And a + c > b And b + c > a Then
        p = (a + b + c) / 2
        Text4.Text = Sqr(p * (p - a) * (p - b) * (p - c))
        Text5.Text = a + b + c
    Else
        MsgBox "不能组成三角形", 0 + 48, "提示"
    End If
End Sub

Private Sub Command2_Click()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
End Sub

Private Sub Command3_Click()
    Unload Form6
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
