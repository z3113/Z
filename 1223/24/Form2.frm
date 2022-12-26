VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "抛硬币"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6615
   LinkTopic       =   "Form2"
   ScaleHeight     =   5085
   ScaleWidth      =   6615
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   720
      Top             =   840
   End
   Begin VB.Frame Frame1 
      Caption         =   "猜一猜"
      Enabled         =   0   'False
      Height          =   1215
      Left            =   1440
      TabIndex        =   1
      Top             =   3240
      Width           =   3735
      Begin VB.OptionButton Option2 
         Caption         =   "反面"
         Height          =   375
         Left            =   2280
         TabIndex        =   3
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "正面"
         Height          =   375
         Left            =   720
         TabIndex        =   2
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "抛硬币"
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   2040
      Width           =   1575
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
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   4680
      TabIndex        =   4
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   1410
      Left            =   2520
      Picture         =   "Form2.frx":0000
      Top             =   240
      Width           =   1500
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
    Label1.Caption = ""
    Frame1.Enabled = True
    Timer1.Enabled = True
    Option1.Value = False
    Option2.Value = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub

Private Sub Option1_Click()
    Timer1.Enabled = False
    Frame1.Enabled = False
    If a = 1 Then
        Label1.ForeColor = vbRed
        Label1.Caption = "恭喜您猜对了！"
    ElseIf a = 2 Then
        Label1.ForeColor = vbBlue
        Label1.Caption = "猜错了！"
    End If
End Sub

Private Sub Option2_Click()
    Timer1.Enabled = False
    Frame1.Enabled = False
    If a = 2 Then
        Label1.ForeColor = vbRed
        Label1.Caption = "恭喜您猜对了！"
    ElseIf a = 1 Then
        Label1.ForeColor = vbBlue
        Label1.Caption = "猜错了！"
    End If
End Sub

Private Sub Timer1_Timer()
    Randomize
    a = Int(Rnd * 2 + 1)
    Select Case a
        Case 1
            Image1.Picture = LoadPicture("zm.gif")
        Case 2
            Image1.Picture = LoadPicture("fm.gif")
    End Select
End Sub
