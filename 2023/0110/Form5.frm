VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00C0C0FF&
   Caption         =   "移动文字"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8175
   LinkTopic       =   "Form5"
   ScaleHeight     =   4365
   ScaleWidth      =   8175
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   5400
      Top             =   3600
   End
   Begin VB.CommandButton Command1 
      Caption         =   "关闭本窗体"
      Height          =   615
      Left            =   6000
      TabIndex        =   0
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "移动方向"
      Height          =   975
      Left            =   840
      TabIndex        =   3
      Top             =   3120
      Width           =   3975
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0C0FF&
         Caption         =   "下"
         Height          =   375
         Left            =   3120
         TabIndex        =   7
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0C0FF&
         Caption         =   "上"
         Height          =   375
         Left            =   2160
         TabIndex        =   6
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "右"
         Height          =   375
         Left            =   1320
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "左"
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "祝你考试顺利"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "移动文字"
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
      Left            =   3120
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Form5
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub

Private Sub Option1_Click()
    Timer1.Enabled = True
End Sub

Private Sub Option2_Click()
    Timer1.Enabled = True
End Sub

Private Sub Option3_Click()
    Timer1.Enabled = True
End Sub

Private Sub Option4_Click()
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    If Option1.Value = True Then
        Label2.Left = Label2.Left - 100
        If Label2.Left <= -Label2.Width Then Label2.Left = Form5.Width
    ElseIf Option2.Value = True Then
        Label2.Left = Label2.Left + 100
        If Label2.Left >= Form5.Width Then Label2.Left = -Label2.Width
    ElseIf Option3.Value = True Then
        Label2.Top = Label2.Top - 100
        If Label2.Top <= -Label2.Height Then Label2.Top = Form5.Height
    ElseIf Option4.Value = True Then
        Label2.Top = Label2.Top + 100
        If Label2.Top >= Form5.Height Then Label2.Top = -Label2.Height
    End If
End Sub
