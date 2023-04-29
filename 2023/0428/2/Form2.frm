VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "基本操作"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7920
   LinkTopic       =   "Form2"
   ScaleHeight     =   3960
   ScaleWidth      =   7920
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
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
      Left            =   5040
      TabIndex        =   10
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   1440
      Width           =   855
   End
   Begin VB.CheckBox Check3 
      Caption         =   "商务套餐30元"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CheckBox Check2 
      Caption         =   "标准套餐23元"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Caption         =   "儿童套餐18元"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "份"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3480
      TabIndex        =   9
      Top             =   2880
      Width           =   210
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "份"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3480
      TabIndex        =   8
      Top             =   2160
      Width           =   210
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "份"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3480
      TabIndex        =   7
      Top             =   1440
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "菜单"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   600
      TabIndex        =   0
      Top             =   1080
      Width           =   480
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
    If Check1.Value = 1 Then Text1.Enabled = True Else Text1.Enabled = False
End Sub

Private Sub Check2_Click()
    If Check2.Value = 1 Then Text1.Enabled = True Else Text1.Enabled = False
End Sub

Private Sub Check3_Click()
    If Check3.Value = 1 Then Text1.Enabled = True Else Text1.Enabled = False
End Sub

Private Sub Command1_Click()
    Dim a%, b%, c%
    If Text1.Enabled = True Then a = Int(Text1.Text)
    If Text2.Enabled = True Then b = Int(Text2.Text)
    If Text3.Enabled = True Then c = Int(Text3.Text)
    MsgBox "一共" & a * 18 + b * 23 + c * 30 & "元", vbOKCancel + 64, "点餐"
End Sub

Private Sub Form_Load()
    Form1.Show
End Sub
