VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "任务四"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5190
   LinkTopic       =   "Form5"
   ScaleHeight     =   4320
   ScaleWidth      =   5190
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "返回（ESC）"
      Height          =   495
      Left            =   3720
      TabIndex        =   11
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清除"
      Height          =   495
      Left            =   3720
      TabIndex        =   10
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "计算价格"
      Height          =   495
      Left            =   3720
      TabIndex        =   9
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "优惠价："
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "原价："
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "数量："
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "单价："
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "本商场所有商品8.9折"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Text3.Text = Val(Text1.Text) * Val(Text2.Text)
    Text4.Text = Val(Text3.Text) * 0.89
End Sub

Private Sub Command2_Click()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
End Sub

Private Sub Command3_Click()
    Unload Form5
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
