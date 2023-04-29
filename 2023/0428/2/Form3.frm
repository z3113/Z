VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00C0FFFF&
   Caption         =   "简单应用"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9750
   LinkTopic       =   "Form3"
   ScaleHeight     =   5685
   ScaleWidth      =   9750
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6240
      Top             =   4080
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "停止"
      Height          =   615
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4080
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "运动"
      Height          =   615
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4080
      Width           =   1695
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   4095
      Left            =   9120
      Max             =   500
      Min             =   10
      TabIndex        =   1
      Top             =   480
      Value           =   100
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Height          =   3255
      Left            =   240
      Picture         =   "Form3.frx":0000
      ScaleHeight     =   3195
      ScaleWidth      =   8595
      TabIndex        =   0
      Top             =   600
      Width           =   8655
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   6960
         TabIndex        =   6
         Top             =   1680
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   960
         TabIndex        =   5
         Top             =   1680
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8280
      TabIndex        =   2
      Top             =   4200
      Width           =   615
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x%

Private Sub Command1_Click()
    Dim a$
    x = 100
    a = InputBox("请输入运动文字（4字以上）", "运动", "改革开放")
    Label2.Caption = Mid(a, 1, 2)
    Label3.Caption = Mid(a, Len(a) - 1, 2)
    Label2.Visible = True
    Label3.Visible = True
    Timer1.Enabled = True
    Command2.Visible = True
    Command1.Visible = False
End Sub

Private Sub Command2_Click()
    Timer1.Enabled = False
    Command1.Visible = True
    Command2.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub

Private Sub Timer1_Timer()
    Label2.Left = Label2.Left + x
    Label3.Left = Label3.Left - x
    If (Label2.Left <= 0 And Label3.Left >= Picture1.Width - Label3.Left) Or Label2.Left + Label2.Width >= Label3.Left Then x = -x
End Sub

Private Sub VScroll1_Change()
    Label1.Caption = VScroll1.Value
    Timer1.Interval = VScroll1.Value
End Sub

Private Sub VScroll1_Scroll()
    Label1.Caption = VScroll1.Value
    Timer1.Interval = VScroll1.Value
End Sub
