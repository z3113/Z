VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00C0E0FF&
   ClientHeight    =   4815
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8085
   LinkTopic       =   "Form5"
   ScaleHeight     =   4815
   ScaleWidth      =   8085
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command5 
      Cancel          =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6600
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5280
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3720
      TabIndex        =   2
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2040
      TabIndex        =   1
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   24.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   600
      TabIndex        =   3
      Top             =   480
      Width           =   6735
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    Unload Form4
    Form1.Show
End Sub

Private Sub Command2_Click()
    Unload Form4
    Form2.Show
End Sub

Private Sub Command3_Click()
    Unload Form4
    Form3.Show
End Sub

Private Sub Command4_Click()
    Unload Form4
    Form5.Show
End Sub

Private Sub Command5_Click()
    End
End Sub

Private Sub Form_Load()

End Sub
