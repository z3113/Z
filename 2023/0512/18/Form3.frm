VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "交换值"
   ClientHeight    =   3375
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8175
   LinkTopic       =   "Form3"
   ScaleHeight     =   3375
   ScaleWidth      =   8175
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "清空"
      Height          =   495
      Left            =   5760
      TabIndex        =   6
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "交换"
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "产生"
      Height          =   495
      Left            =   960
      TabIndex        =   4
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   1440
      Width           =   6375
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   6375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "交换前："
      Height          =   180
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "交换前："
      Height          =   180
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   720
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim a(25) As String, i As Integer

Private Sub Command1_Click()
    Randomize
    Text1.Text = ""
    For i = 1 To 25
        a(i) = Chr(Int(Rnd * 26) + 65)
        Text1.Text = Text1.Text & a(i) & " "
    Next i
End Sub

Private Sub Command2_Click()
    Text2.Text = ""
'    For i = 25 To 1 Step -1
'        If i = 13 Then
'            Text2.Text = Text2.Text & a(i) & " "
'        Else
'            Text2.Text = Text2.Text & LCase(a(i)) & " "
'        End If
'    Next i
    For i = 1 To 12
        a(0) = LCase(a(i))
        a(i) = LCase(a(26 - i))
        a(26 - i) = a(0)
    Next i
    For i = 1 To 25
        Text2.Text = Text2.Text & a(i) & " "
    Next i
End Sub

Private Sub Command3_Click()
    Text1.Text = ""
    Text2.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
