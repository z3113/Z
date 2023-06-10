VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FFC0C0&
   Caption         =   "迷你文本编辑器"
   ClientHeight    =   5295
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8535
   LinkTopic       =   "Form3"
   ScaleHeight     =   5295
   ScaleWidth      =   8535
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "返回"
      Height          =   495
      Left            =   5640
      TabIndex        =   11
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "样式"
      Height          =   975
      Left            =   840
      TabIndex        =   7
      Top             =   3840
      Width           =   3855
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFC0FF&
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "U"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "颜色"
      Height          =   855
      Left            =   4920
      TabIndex        =   5
      Top             =   3240
      Width           =   2655
      Begin VB.ComboBox Combo3 
         Enabled         =   0   'False
         Height          =   300
         ItemData        =   "Form3.frx":0000
         Left            =   240
         List            =   "Form3.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "字号"
      Height          =   1335
      Left            =   4920
      TabIndex        =   3
      Top             =   1680
      Width           =   2655
      Begin VB.ComboBox Combo2 
         Enabled         =   0   'False
         Height          =   720
         ItemData        =   "Form3.frx":0023
         Left            =   240
         List            =   "Form3.frx":0036
         Style           =   1  'Simple Combo
         TabIndex        =   4
         Text            =   "请选择字号"
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "字体"
      Height          =   1215
      Left            =   4920
      TabIndex        =   1
      Top             =   240
      Width           =   2655
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   300
         ItemData        =   "Form3.frx":004E
         Left            =   240
         List            =   "Form3.frx":005E
         TabIndex        =   2
         Text            =   "请选择字体"
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.TextBox Text1 
      Height          =   3255
      Left            =   840
      TabIndex        =   0
      Text            =   "请输入内容"
      Top             =   360
      Width           =   3855
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Combo1_Change()
    Text1.FontName = Combo1.Text
End Sub

Private Sub Combo1_Click()
    Text1.FontName = Combo1.Text
End Sub

Private Sub Combo2_Change()
    Text1.FontSize = Combo2.Text
End Sub

Private Sub Combo2_Click()
    Text1.FontSize = Combo2.Text
End Sub

Private Sub Combo3_Click()
    Select Case Combo3.ListIndex
        Case 0
            Text1.ForeColor = vbRed
        Case 1
            Text1.ForeColor = vbBlue
        Case 2
            Text1.ForeColor = vbGreen
    End Select
End Sub

Private Sub Command1_Click()
    Text1.FontUnderline = Not Text1.FontUnderline
End Sub

Private Sub Command2_Click()
    Text1.FontItalic = Not Text1.FontItalic
End Sub

Private Sub Command3_Click()
    Text1.FontBold = Not Text1.FontBold
End Sub

Private Sub Command4_Click()
    Unload Form3
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub

Private Sub Text1_Change()
    Combo1.Enabled = True
    Combo2.Enabled = True
    Combo3.Enabled = True
End Sub
