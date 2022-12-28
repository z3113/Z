VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7680
   LinkTopic       =   "Form2"
   ScaleHeight     =   4245
   ScaleWidth      =   7680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "停止移动（&Z）"
      Height          =   495
      Left            =   2520
      TabIndex        =   13
      Top             =   3600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始移动（&S）"
      Height          =   495
      Left            =   2520
      TabIndex        =   12
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "文字效果"
      Height          =   1335
      Left            =   6360
      TabIndex        =   8
      Top             =   1680
      Width           =   1215
      Begin VB.CheckBox Check3 
         Caption         =   "下划线"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   855
      End
      Begin VB.CheckBox Check2 
         Caption         =   "删除线"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "加粗"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Value           =   1  'Checked
         Width           =   735
      End
   End
   Begin VB.Frame zitiyanse 
      Caption         =   "字体验色"
      Height          =   1335
      Left            =   6360
      TabIndex        =   4
      Top             =   240
      Width           =   1215
      Begin VB.Label Label4 
         BackColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   6360
      TabIndex        =   3
      Text            =   "500"
      Top             =   3120
      Width           =   615
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      LargeChange     =   10
      Left            =   240
      Max             =   10
      Min             =   500
      TabIndex        =   2
      Top             =   3120
      Value           =   500
      Width           =   5895
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   360
      Top             =   360
   End
   Begin VB.PictureBox Picture1 
      Height          =   2775
      Left            =   240
      ScaleHeight     =   2715
      ScaleWidth      =   5835
      TabIndex        =   0
      Top             =   240
      Width           =   5895
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   720
         Top             =   120
      End
      Begin VB.Label Label1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   0
         TabIndex        =   1
         Top             =   1080
         Width           =   3135
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim a%

Private Sub Check1_Click()
    Label1.FontBold = Not Label1.FontBold
End Sub

Private Sub Check2_Click()
    Label1.FontStrikethru = Not Label1.FontStrikethru
End Sub

Private Sub Check3_Click()
    Label1.FontUnderline = Not Label1.FontUnderline
End Sub

Private Sub Command1_Click()
    Command1.Visible = False
    Command2.Visible = True
    Timer2.Enabled = True
End Sub

Private Sub Command2_Click()
    Command1.Visible = True
    Command2.Visible = False
    Timer2.Enabled = False
End Sub

Private Sub Form_Activate()
    Label1.Caption = Now
    a = 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub

Private Sub HScroll1_Change()
    Text1.Text = HScroll1.Value
    Timer2.Interval = HScroll1.Value
End Sub

Private Sub HScroll1_Scroll()
    Text1.Text = HScroll1.Value
    Timer2.Interval = HScroll1.Value
End Sub

Private Sub Label2_Click()
    Label1.ForeColor = vbRed
End Sub

Private Sub Label3_Click()
    Label1.ForeColor = vbYellow
End Sub

Private Sub Label4_Click()
    Label1.ForeColor = vbBlue
End Sub

Private Sub Timer1_Timer()
    Label1.Caption = Now
End Sub

Private Sub Timer2_Timer()
    Label1.Left = Label1.Left + a
    If Label1.Left >= Picture1.Width - Label1.Left Then
        a = -100
    ElseIf Label1.Left <= 0 Then
        a = 100
    End If
End Sub
