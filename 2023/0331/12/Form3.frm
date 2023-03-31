VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "水仙花"
   ClientHeight    =   4050
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5895
   LinkTopic       =   "Form3"
   ScaleHeight     =   4050
   ScaleWidth      =   5895
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "返回"
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "显示水仙花数"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   3240
      Width           =   1095
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim i%
    Cls
    Print "三位数的水仙花数有："
    For i = 100 To 999
        If (i \ 100) ^ 3 + (i \ 10 Mod 10) ^ 3 + (i Mod 10) ^ 3 = i Then Print i;
    Next i
End Sub

Private Sub Command2_Click()
    Unload Form3
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form2.Show
End Sub
