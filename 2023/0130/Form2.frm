VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Font Sitting"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6135
   LinkTopic       =   "Form2"
   ScaleHeight     =   4590
   ScaleWidth      =   6135
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame3 
      Caption         =   "Style"
      Height          =   1935
      Left            =   4320
      TabIndex        =   13
      Top             =   2520
      Width           =   1575
      Begin VB.CheckBox Check3 
         Caption         =   "Underline"
         Height          =   375
         Left            =   360
         TabIndex        =   16
         Top             =   1320
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Incline"
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Top             =   840
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Bold"
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Color"
      Height          =   1935
      Left            =   2280
      TabIndex        =   9
      Top             =   2520
      Width           =   1575
      Begin VB.OptionButton Option6 
         Caption         =   "Red"
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Blue"
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   840
         Width           =   735
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Green"
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   1320
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Font"
      Height          =   1935
      Left            =   240
      TabIndex        =   5
      Top             =   2520
      Width           =   1575
      Begin VB.OptionButton Option3 
         Caption         =   "楷体"
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   1320
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         Caption         =   "隶书"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   840
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "宋体"
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "unload"
      Height          =   495
      Left            =   4800
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "reset"
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   2
      Text            =   "10"
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Font Size"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
    Text1.FontBold = Not Text1.FontBold
End Sub

Private Sub Check2_Click()
    Text1.FontItalic = Not Text1.FontItalic
End Sub

Private Sub Check3_Click()
    Text1.FontUnderline = Not Text1.FontUnderline
End Sub

Private Sub Command1_Click()
    Option1.Value = False
    Option2.Value = False
    Option3.Value = False
    Option4.Value = False
    Option5.Value = False
    Option6.Value = False
    Check1.Value = 0
    Check2.Value = 0
    Check3.Value = 0
    Text2.Text = 10
    Text1.FontName = "宋体"
    Text1.ForeColor = vbBlack
End Sub

Private Sub Command2_Click()
    Unload Form2
End Sub

Private Sub Option1_Click()
    Text1.FontName = "宋体"
End Sub

Private Sub Option2_Click()
    Text1.FontName = "隶书"
End Sub

Private Sub Option3_Click()
    Text1.FontName = "楷体"
End Sub

Private Sub Option4_Click()
    Text1.ForeColor = vbGreen
End Sub

Private Sub Option5_Click()
    Text1.ForeColor = vbBlue
End Sub

Private Sub Option6_Click()
    Text1.ForeColor = vbRed
End Sub

Private Sub Text2_Change()
    Text1.FontSize = Val(Text2.Text)
End Sub
