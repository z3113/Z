VERSION 5.00
Begin VB.Form Form11 
   Caption         =   "²ÂµÆÃÕ"
   ClientHeight    =   3720
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6165
   LinkTopic       =   "Form11"
   ScaleHeight     =   3720
   ScaleWidth      =   6165
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command1 
      Caption         =   "´ð°¸"
      Height          =   495
      Left            =   4440
      TabIndex        =   1
      Top             =   3000
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   2460
      Left            =   4080
      Picture         =   "Form11.frx":0000
      ScaleHeight     =   2400
      ScaleWidth      =   1755
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Cls
    Dim i%, j%, k%, l%, a
    For i = 1 To 9
        For j = 0 To 9
            For k = 1 To 9
                For l = 0 To 9
                    If (i & j & k & l) - (k & l & k) = (i & j & k) Then Print i; j; k; l
                Next l
            Next k
        Next j
    Next i
End Sub
