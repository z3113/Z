VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "代码调试1"
   ClientHeight    =   3375
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5295
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   15
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF8080&
   LinkTopic       =   "Form2"
   ScaleHeight     =   3375
   ScaleWidth      =   5295
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "代码2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "代码1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代码1结果显示在这里"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   2865
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim a(10) As Integer, num As Integer, i As Integer
    num = -2
    For i = 1 To 5
        a(i) = i
        num = num + a(i)
    Next i
    Label1.Caption = num
End Sub

Private Sub Command2_Click()
    Cls
    Dim a(10) As Integer, b(5) As Integer, n As Integer, i As Integer
    n = 5
    For i = 1 To 5
        a(i) = i
        b(n) = 2 * i + a(i)
    Next i
    Print a(i); b(n)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
