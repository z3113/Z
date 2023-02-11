VERSION 5.00
Begin VB.Form Form3 
   ClientHeight    =   4890
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7095
   LinkTopic       =   "Form4"
   ScaleHeight     =   4890
   ScaleWidth      =   7095
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command6 
      Cancel          =   -1  'True
      Caption         =   "返回"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   9
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "四舍五入保留4位小数"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4800
      TabIndex        =   7
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "保留4位小数"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3480
      TabIndex        =   6
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "四舍五入保留3位小数"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1440
      TabIndex        =   5
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "保留2位小数"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   4
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   1320
      TabIndex        =   3
      Top             =   2280
      Width           =   5055
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   4440
      TabIndex        =   2
      Text            =   "7"
      Top             =   0
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "÷"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Text            =   "50"
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   1800
      TabIndex        =   8
      Top             =   1080
      Width           =   3855
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a#

Private Sub Command1_Click()
    a = Text1.Text / Text2.Text
    Label1.Caption = a
End Sub

Private Sub Command2_Click()
    Text3.Text = Int(a * 100) / 100
End Sub

Private Sub Command3_Click()
    Text3.Text = Round(a * 1000) / 1000
End Sub

Private Sub Command4_Click()
    Text3.Text = Int(a * 10000) / 10000
End Sub

Private Sub Command5_Click()
    Text3.Text = Int(a * 10000 + 0.5) / 10000
End Sub

Private Sub Command6_Click()
    Unload Form3
    Form4.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Form3
    Form4.Show
End Sub
