VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "倒计时"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   4230
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2160
      Top             =   3000
   End
   Begin VB.CommandButton Command2 
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
      Left            =   3000
      TabIndex        =   3
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始倒计时"
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
      Left            =   360
      TabIndex        =   2
      Top             =   2880
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
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   4215
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a%
Option Explicit

Private Sub Command1_Click()
    a = 5
    Label1.Caption = "还有" & a & "秒"
    Text1.Text = Hour(Time) & ":" & Minute(Time) & ":" & Second(Time)
    Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
    Unload Form2
    Form1.Show
End Sub

Private Sub Timer1_Timer()
    a = a - 1
    Label1.Caption = "还有" & a & "秒"
    Text1.Text = Hour(Time) & ":" & Minute(Time) & ":" & Second(Time)
    If a = 0 Then
        Timer1.Enabled = False
        MsgBox "时间到", 1 + 16, "倒计时"
    End If
End Sub
