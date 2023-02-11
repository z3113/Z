VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "倒计时"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4710
   LinkTopic       =   "Form2"
   ScaleHeight     =   4215
   ScaleWidth      =   4710
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2160
      Top             =   2400
   End
   Begin VB.CommandButton Command2 
      Caption         =   "返回"
      Height          =   615
      Left            =   2760
      TabIndex        =   3
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始倒计时"
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1200
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a%

Private Sub Command1_Click()
    a = 10
    Text1.Text = Time
    Label1.Caption = "还有" & a & "秒"
    Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
    Unload Form2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub

Private Sub Timer1_Timer()
    a = a - 1
    Text1.Text = Time
    Label1.Caption = "还有" & a & "秒"
    If a = 0 Then
        MsgBox "时间到", vbOKCancel + 16, "倒计时"
        Timer1.Enabled = False
    End If
End Sub
