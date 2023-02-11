VERSION 5.00
Begin VB.Form Form5 
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form5"
   ScaleHeight     =   3030
   ScaleWidth      =   4680
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command3 
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim n%
Private Sub Command1_Click()
    Const MY! = -9 + 3.7
    Const e = 2.718281828
    Const a1 = 4
    Const b = e + 1
    Const KH As String = ""
    Const Pi = 3.1415926
    Print MY, e, b, KH, Pi
End Sub

Private Sub Command2_Click()
    Dim sum%
    n = n + 1
    sum = sum + 1
    Print " "; n; sum
End Sub

Private Sub Command3_Click()
    Dim x%
    Static y%
    x = x + 1
    y = y + 1
    Print ""; x; y
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Form5
    Form1.Show
End Sub
