VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "5单选框和IF"
   ClientHeight    =   3540
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5400
   LinkTopic       =   "Form6"
   ScaleHeight     =   3540
   ScaleWidth      =   5400
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "清空"
      Height          =   495
      Left            =   3840
      TabIndex        =   7
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "显示"
      Height          =   495
      Left            =   3840
      TabIndex        =   6
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   1215
      Left            =   600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1920
      Width           =   2775
   End
   Begin VB.OptionButton Option2 
      Caption         =   "女"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   4
      Top             =   1200
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "男"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   1200
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "性别"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "姓名"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim a$
    If Option1.Value = True Then
        a = "男"
    Else
        a = "女"
    End If
    Text2.Text = "姓名：" & Text1.Text & vbCrLf & "性别：" & a
End Sub

Private Sub Command2_Click()
    Text1.Text = ""
    Text2.Text = ""
    Option1.Value = False
    Option2.Value = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
