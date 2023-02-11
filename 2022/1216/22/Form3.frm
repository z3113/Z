VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00C0E0FF&
   Caption         =   "任务二：滚动字幕"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6375
   LinkTopic       =   "Form3"
   ScaleHeight     =   4245
   ScaleWidth      =   6375
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "滚动方向"
      Height          =   735
      Left            =   2520
      TabIndex        =   6
      Top             =   3240
      Width           =   3615
      Begin VB.HScrollBar HScroll1 
         Height          =   375
         LargeChange     =   3
         Left            =   120
         Max             =   2000
         TabIndex        =   11
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2400
         TabIndex        =   12
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "滚动方向"
      Height          =   735
      Left            =   2520
      TabIndex        =   5
      Top             =   2160
      Width           =   3615
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "右"
         Height          =   375
         Left            =   2760
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "左"
         Height          =   375
         Left            =   1920
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "下"
         Height          =   375
         Left            =   1080
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "上"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5040
      Top             =   1080
   End
   Begin VB.CommandButton Command4 
      Caption         =   "返回"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "暂停"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "开始滚动"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "设置滚动内容"
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   6375
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
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   2895
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a%
Option Explicit

Private Sub Command1_Click()
    Label1.Caption = InputBox("请输入滚动字幕内容", "字幕内容", "我要好好学习，天天向上")
End Sub

Private Sub Command2_Click()
    Timer1.Enabled = True
End Sub

Private Sub Command3_Click()
    Timer1.Enabled = False
End Sub

Private Sub Command4_Click()
    Unload Form3
End Sub

Private Sub Form_Activate()
    Label1.Caption = InputBox("请输入滚动字幕内容", "字幕内容", "我要好好学习，天天向上")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub

Private Sub HScroll1_Change()
    a = Val(HScroll1.Value)
    Label2.Caption = HScroll1.Value
End Sub

Private Sub HScroll1_Scroll()
    a = Val(HScroll1.Value)
    Label2.Caption = HScroll1.Value
End Sub

Private Sub Timer1_Timer()
    If Option1.Value = True Then
        Label1.Top = Label1.Top - a
        If Label1.Top <= -Label1.Height Then Label1.Top = Form3.Height
    ElseIf Option2.Value = True Then
        Label1.Top = Label1.Top + a
        If Label1.Top >= Form3.Height Then Label1.Top = -Label1.Height
    ElseIf Option3.Value = True Then
        Label1.Left = Label1.Left - a
        If Label1.Left <= Form3.Width Then Label1.Left = Form3.Height
    ElseIf Option4.Value = True Then
        Label1.Left = Label1.Left + a
        If Label1.Left >= Form3.Width Then Label1.Left = -Label1.Width
    End If
End Sub
