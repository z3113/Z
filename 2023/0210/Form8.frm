VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00C0C0FF&
   Caption         =   "移动文字"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7830
   LinkTopic       =   "Form8"
   ScaleHeight     =   4350
   ScaleWidth      =   7830
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   7200
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "关闭本窗体"
      Height          =   615
      Left            =   5640
      TabIndex        =   0
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "移动方向"
      Height          =   855
      Left            =   600
      TabIndex        =   3
      Top             =   3120
      Width           =   4695
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0C0FF&
         Caption         =   "下"
         Height          =   255
         Left            =   3720
         TabIndex        =   7
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0C0FF&
         Caption         =   "上"
         Height          =   255
         Left            =   2640
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "右"
         Height          =   255
         Left            =   1560
         TabIndex        =   5
         Top             =   360
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "左"
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "祝你考试顺利"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   2340
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
      Left            =   3000
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Form8
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
        If Label2.Left <= -Label2.Width Then Label2.Left = Form8.Width
    ElseIf Option2.Value = True Then
        Label2.Left = Label2.Left + 100
        If Label2.Left >= Form8.Width Then Label2.Left = -Label2.Width
    ElseIf Option3.Value = True Then
        Label2.Top = Label2.Top - 100
        If Label2.Top <= -Label2.Height Then Label2.Top = Form8.Height
    ElseIf Option4.Value = True Then
        Label2.Top = Label2.Top + 100
        If Label2.Top >= Form8.Height Then Label2.Top = -Label2.Height
    End If
End Sub
