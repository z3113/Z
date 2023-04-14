VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FFC0C0&
   Caption         =   "素数算法"
   ClientHeight    =   6735
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7590
   LinkTopic       =   "Form3"
   ScaleHeight     =   6735
   ScaleWidth      =   7590
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "返回"
      Height          =   495
      Left            =   5520
      TabIndex        =   8
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清空"
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "产生"
      Height          =   495
      Left            =   720
      TabIndex        =   6
      Top             =   5880
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   4335
      Left            =   5520
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   4335
      Left            =   2880
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   4335
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "水仙花数："
      Height          =   180
      Left            =   5640
      TabIndex        =   4
      Top             =   240
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "素数："
      Height          =   180
      Left            =   3000
      TabIndex        =   2
      Top             =   240
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "随机产生的数："
      Height          =   180
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1260
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Randomize
    Dim i%, j%, a%
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    For i = 1 To 50
        a = Int(Rnd * 900 + 1)
        Text1.Text = Text1.Text & a & vbCrLf
        For j = 2 To a
            If a Mod j = 0 Then Exit For
        Next j
        If j = a Then Text2.Text = Text2.Text & a & vbCrLf
        If (100 <= a And a <= 999) And ((a \ 100) ^ 3 + (a \ 10 Mod 10) ^ 3 + (a Mod 10) ^ 3 = a) Then Text3.Text = Text3.Text & a & vbCrLf
    Next i
End Sub

Private Sub Command2_Click()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
End Sub

Private Sub Command3_Click()
    Unload Form3
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
