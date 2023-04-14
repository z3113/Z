VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00C0C0FF&
   Caption         =   "数列计算"
   ClientHeight    =   6495
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8655
   LinkTopic       =   "Form5"
   ScaleHeight     =   6495
   ScaleWidth      =   8655
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "清空"
      Height          =   495
      Left            =   6120
      TabIndex        =   5
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "计算（方法二）"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3480
      TabIndex        =   4
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "计算（方法一）"
      Enabled         =   0   'False
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   5520
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   120
      ScaleHeight     =   3675
      ScaleWidth      =   7995
      TabIndex        =   2
      Top             =   1200
      Width           =   8055
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      LargeChange     =   3
      Left            =   240
      Max             =   50
      Min             =   3
      TabIndex        =   0
      Top             =   360
      Value           =   3
      Width           =   2295
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2640
      TabIndex        =   6
      Top             =   360
      Width           =   120
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "数列：前一个数的三倍减去在前面一个数。1  1 2 5 13 ......"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      Top             =   360
      Width           =   4695
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Picture1.Cls
    Dim i%, a#, b#, c#
    a = 1
    b = 1
    Picture1.Print a; b;
    For i = 3 To Val(HScroll1.Value)
        c = 3 * b - a
        Picture1.Print c;
        a = b
        b = c
        If i Mod 6 = 0 Then Picture1.Print
    Next i
End Sub

Private Sub Command2_Click()
    Dim a#, b#, i%
    Picture1.Cls
    a = 1
    b = 1
    Picture1.Print a; b;
    For i = 2 To Val(HScroll1.Value) / 2
        a = 3 * b - a
        b = 3 * a - b
        Picture1.Print a; b;
        If i Mod 3 = 0 Then Picture1.Print
    Next i
End Sub

Private Sub Command3_Click()
    Picture1.Cls
    HScroll1.Value = 3
    Command1.Enabled = False
    Command2.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub

Private Sub HScroll1_Change()
    If Val(HScroll1.Value) Mod 2 <> 0 Then
        MsgBox "请重新设置滚动条的值，需要为偶数", vbOKOnly + 16, "出错！"
    Else
        Label2.Caption = HScroll1.Value
    End If
    Command1.Enabled = True
    Command2.Enabled = True
End Sub

Private Sub HScroll1_Scroll()
    If Val(HScroll1.Value) Mod 2 <> 0 Then
        MsgBox "请重新设置滚动条的值，需要为偶数", vbOKOnly + 16, "出错！"
    Else
        Label2.Caption = HScroll1.Value
    End If
    Command1.Enabled = True
    Command2.Enabled = True
End Sub
