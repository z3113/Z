VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "倒计时"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4695
   LinkTopic       =   "Form2"
   ScaleHeight     =   4245
   ScaleWidth      =   4695
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2160
      Top             =   3120
   End
   Begin VB.CommandButton Command2 
      Caption         =   "返回"
      Height          =   615
      Left            =   2520
      TabIndex        =   3
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始倒计时"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   36
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   36
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a%

Private Sub Command1_Click()
    a = 10
    Text1.Text = Hour(Now) & ":" & Minute(Now) & ":" & Second(Now)
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
    Text1.Text = Hour(Now) & ":" & Minute(Now) & ":" & Second(Now)
    Label1.Caption = "还有" & a & "秒"
    If a = 0 Then
        Timer1.Enabled = False
        MsgBox "时间到", 1 + 16, "倒计时"
    End If
End Sub
