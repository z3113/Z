VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "四则运算"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8220
   LinkTopic       =   "Form4"
   ScaleHeight     =   4215
   ScaleWidth      =   8220
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "随机产生数"
      Height          =   615
      Left            =   6720
      TabIndex        =   11
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "运算符"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   4680
      TabIndex        =   6
      Top             =   360
      Width           =   1935
      Begin VB.OptionButton Option4 
         Caption         =   "／"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   26.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   10
         Top             =   2760
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         Caption         =   "×"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   26.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   9
         Top             =   2040
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "－"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   26.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   8
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "＋"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   26.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   7
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2160
      TabIndex        =   5
      Top             =   3000
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2160
      TabIndex        =   3
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2160
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "结   果"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "运算数2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "运算数1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Randomize
    Text1.Text = Int(Rnd * 10)
    Text2.Text = Int(Rnd * 10)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub

Private Sub Option1_Click()
    Text3.Text = Val(Text1.Text) + Val(Text2.Text)
End Sub

Private Sub Option2_Click()
    Text3.Text = Val(Text1.Text) - Val(Text2.Text)
End Sub

Private Sub Option3_Click()
    Text3.Text = Val(Text1.Text) * Val(Text2.Text)
End Sub

Private Sub Option4_Click()
    Text3.Text = Val(Text1.Text) / Val(Text2.Text)
End Sub
