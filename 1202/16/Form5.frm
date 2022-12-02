VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "任务四：字符转换"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7575
   LinkTopic       =   "Form5"
   ScaleHeight     =   6750
   ScaleWidth      =   7575
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command5 
      Cancel          =   -1  'True
      Caption         =   "返回（ESC）"
      Height          =   495
      Left            =   5400
      TabIndex        =   12
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "大小写字母相互转换"
      Height          =   2655
      Left            =   240
      TabIndex        =   7
      Top             =   2640
      Width           =   7095
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   240
         TabIndex        =   11
         Top             =   1800
         Width           =   6495
      End
      Begin VB.CommandButton Command4 
         Caption         =   "大写改小写"
         Height          =   615
         Left            =   4080
         TabIndex        =   10
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "小写改大写"
         Height          =   615
         Left            =   1320
         TabIndex        =   9
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
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
         Height          =   615
         Left            =   240
         TabIndex        =   8
         Text            =   "abcdefghijkl1234567890ABCDEFGHIJK"
         Top             =   360
         Width           =   6495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ASCII值和字符相互转换"
      Height          =   2055
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7095
      Begin VB.CommandButton Command2 
         Caption         =   "将左边ASC值转换成字符"
         Height          =   495
         Left            =   1680
         TabIndex        =   6
         Top             =   1200
         Width           =   2895
      End
      Begin VB.CommandButton Command1 
         Caption         =   "将左边字母转换成ASC值"
         Height          =   495
         Left            =   1680
         TabIndex        =   5
         Top             =   480
         Width           =   2895
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5160
         TabIndex        =   4
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Text            =   "65"
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5160
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   1
         Text            =   "A"
         Top             =   480
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Text2.Text = Asc(Text1.Text)
End Sub

Private Sub Command2_Click()
    Text4.Text = Chr(Text3.Text)
End Sub

Private Sub Command3_Click()
    Text6.Text = LCase(Text5.Text)
End Sub

Private Sub Command4_Click()
    Text6.Text = UCase(Text5.Text)
End Sub

Private Sub Command5_Click()
    Unload Form5
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
