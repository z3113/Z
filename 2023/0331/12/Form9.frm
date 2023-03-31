VERSION 5.00
Begin VB.Form Form9 
   Caption         =   "勾股定理"
   ClientHeight    =   4575
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7935
   LinkTopic       =   "Form9"
   ScaleHeight     =   4575
   ScaleWidth      =   7935
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "答案"
      Height          =   615
      Left            =   5520
      TabIndex        =   1
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   $"Form9.frx":0000
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5280
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Cls
    Print "A", "B", "C"
    Dim i%, j%, k%
    For i = 1 To 30
        For j = 1 To 30
            For k = 1 To 30
                If i ^ 2 + j ^ 2 = k ^ 2 Then Print i, j, k
            Next k
        Next j
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form4.Show
End Sub
