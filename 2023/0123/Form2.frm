VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "基本操作"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7350
   LinkTopic       =   "Form2"
   ScaleHeight     =   4125
   ScaleWidth      =   7350
   StartUpPosition =   3  '窗口缺省
   Begin VB.VScrollBar VScroll1 
      Height          =   1815
      LargeChange     =   5
      Left            =   6360
      Max             =   50
      Min             =   5
      TabIndex        =   14
      Top             =   2040
      Value           =   5
      Width           =   495
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "文字效果"
      Height          =   1935
      Left            =   4560
      TabIndex        =   10
      Top             =   1920
      Width           =   1455
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "下划线"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   1320
         Width           =   975
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "倾斜"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "加粗"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "字体"
      Height          =   1935
      Left            =   2640
      TabIndex        =   6
      Top             =   1920
      Width           =   1455
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "楷体"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "黑体"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "宋体"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "窗体变化"
      Height          =   1815
      Left            =   240
      TabIndex        =   1
      Top             =   2040
      Width           =   1935
      Begin VB.CommandButton Command4 
         Caption         =   "下移"
         Height          =   375
         Left            =   600
         TabIndex        =   5
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "右移"
         Height          =   375
         Left            =   1080
         TabIndex        =   4
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "左移"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "上移"
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   8.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form2.frx":0000
      Top             =   240
      Width           =   5655
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   6960
      TabIndex        =   17
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "50"
      Height          =   255
      Left            =   6960
      TabIndex        =   16
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "5"
      Height          =   255
      Left            =   6960
      TabIndex        =   15
      Top             =   2040
      Width           =   255
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
    Form2.Top = Form2.Top - 200
End Sub

Private Sub Command2_Click()
    Form2.Left = Form2.Left - 200
End Sub

Private Sub Command3_Click()
    Form2.Left = Form2.Left + 200
End Sub

Private Sub Command4_Click()
    Form2.Top = Form2.Top + 200
End Sub

Private Sub Form_DblClick()
    Unload Form2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub

Private Sub Option1_Click()
    Text1.FontName = "宋体"
End Sub

Private Sub Option2_Click()
    Text1.FontName = "黑体"
End Sub

Private Sub Option3_Click()
    Text1.FontName = "楷体"
End Sub

Private Sub Text1_Change()
    Randomize
    Text1.BackColor = RGB(Int(Rnd * 256), Int(Rnd * 256), Int(Rnd * 256))
End Sub

Private Sub VScroll1_Change()
    Label3.Caption = VScroll1.Value
    Text1.FontSize = VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
    Label3.Caption = VScroll1.Value
    Text1.FontSize = VScroll1.Value
End Sub
