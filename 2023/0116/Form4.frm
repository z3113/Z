VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00C0E0FF&
   Caption         =   "综合应用"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6750
   LinkTopic       =   "Form4"
   ScaleHeight     =   4470
   ScaleWidth      =   6750
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5640
      Top             =   1080
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "滚动速度"
      Height          =   975
      Left            =   2280
      TabIndex        =   10
      Top             =   3240
      Width           =   4215
      Begin VB.HScrollBar HScroll1 
         Height          =   375
         LargeChange     =   3
         Left            =   240
         Max             =   2000
         TabIndex        =   11
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2880
         TabIndex        =   12
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "滚动方向"
      Height          =   855
      Left            =   2280
      TabIndex        =   5
      Top             =   1920
      Width           =   4215
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "右"
         Height          =   255
         Left            =   3240
         TabIndex        =   9
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "左"
         Height          =   255
         Left            =   2280
         TabIndex        =   8
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "下"
         Height          =   255
         Left            =   1320
         TabIndex        =   7
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "上"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "返回"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "暂停"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "开始滚动"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "设置滚动内容"
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   6735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2895
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a%

Private Sub Command1_Click()
    Label1.Caption = InputBox("请输入滚动字幕内容", "字母内容", "我要好好学习，天天向上")
End Sub

Private Sub Command2_Click()
    Timer1.Enabled = True
End Sub

Private Sub Command3_Click()
    Timer1.Enabled = False
End Sub

Private Sub Command4_Click()
    Unload Form4
End Sub

Private Sub Form_Activate()
    Label1.Caption = InputBox("请输入滚动字幕内容", "字母内容", "我要好好学习，天天向上")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub

Private Sub HScroll1_Change()
    Label2.Caption = HScroll1.Value
    a = HScroll1.Value
End Sub

Private Sub HScroll1_Scroll()
    Label2.Caption = HScroll1.Value
    a = HScroll1.Value
End Sub

Private Sub Timer1_Timer()
    If Option1.Value = True Then
        Label1.Top = Label1.Top - a
        If Label1.Top <= -Label1.Height Then Label1.Top = Form4.Height
    ElseIf Option2.Value = True Then
        Label1.Top = Label1.Top + a
        If Label1.Top >= Form4.Height Then Label1.Top = -Label1.Height
    ElseIf Option3.Value = True Then
        Label1.Left = Label1.Left - a
        If Label1.Left <= -Label1.Width Then Label1.Left = Form4.Width
    ElseIf Option4.Value = True Then
        Label1.Left = Label1.Left + a
        If Label1.Left >= Form4.Width Then Label1.Left = -Label1.Width
    ElseIf Option1.Value = False And Option2.Value = False And Option3.Value = False And Option4.Value = False Then
        Label1.Left = Label1.Left + a
        If Label1.Left >= Form4.Width Then Label1.Left = -Label1.Width
    End If
End Sub
