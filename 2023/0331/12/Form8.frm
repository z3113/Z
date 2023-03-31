VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "找对错题"
   ClientHeight    =   4110
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7125
   LinkTopic       =   "Form8"
   ScaleHeight     =   4110
   ScaleWidth      =   7125
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "答案"
      Height          =   615
      Left            =   4800
      TabIndex        =   1
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "2、某试卷由26题构成，答对一题得8分，答错1题扣5分。有位考生做完所有考题但得0分，问他做对多少题，做错多少题？"
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
      Left            =   4200
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Cls
    Print "对", "错"
    Dim i%
    For i = 1 To 26
        If i * 8 - (26 - i) * 5 = 0 Then Print i, (26 - i)
    Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form4.Show
End Sub
