VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "意见簿"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7560
   LinkTopic       =   "Form3"
   ScaleHeight     =   4815
   ScaleWidth      =   7560
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame3 
      Caption         =   "字形"
      Height          =   1455
      Left            =   240
      TabIndex        =   6
      Top             =   3120
      Width           =   4335
      Begin VB.CheckBox Check3 
         Caption         =   "斜体"
         Height          =   615
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "粗体"
         Height          =   615
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "下划线"
         Height          =   615
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "退出"
      Height          =   495
      Left            =   6240
      TabIndex        =   5
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "重写"
      Height          =   495
      Left            =   4920
      TabIndex        =   4
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "字号"
      Height          =   1215
      Left            =   5400
      TabIndex        =   3
      Top             =   2520
      Width           =   1695
      Begin VB.OptionButton Option2 
         Caption         =   "12点"
         Height          =   375
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "20点"
         Height          =   375
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "字体"
      Height          =   1935
      Left            =   5400
      TabIndex        =   2
      Top             =   360
      Width           =   1695
      Begin VB.OptionButton Option5 
         Caption         =   "楷体"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         Caption         =   "黑体"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         Caption         =   "宋体"
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.TextBox Text1 
      Height          =   2295
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form3.frx":0000
      Top             =   600
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "请留下宝贵建议："
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
    Text1.FontUnderline = Not Text1.FontUnderline
End Sub

Private Sub Check2_Click()
    Text1.FontBold = Not Text1.FontBold
End Sub

Private Sub Check3_Click()
    Text1.FontItalic = Not Text1.FontItalic
End Sub

Private Sub Command6_Click()
    Text1.Text = ""
End Sub

Private Sub Command7_Click()
    Unload Form3
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub

Private Sub Option1_Click()
    Text1.FontSize = 20
End Sub

Private Sub Option2_Click()
    Text1.FontSize = 12
End Sub

Private Sub Option3_Click()
    Text1.FontName = "宋体"
End Sub

Private Sub Option4_Click()
    Text1.FontName = "黑体"
End Sub

Private Sub Option5_Click()
    Text1.FontName = "楷体"
End Sub
