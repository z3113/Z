VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "测试"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    Randomize
    Dim i%, a%
    a = Int(Rnd * 900) + 100
    For i = 2 To a
        If a Mod i = 0 Then Exit For
    Next i
    If a = i Then
        Print "随机数"; a; "是素数"
    Else
        Print "随机数"; a; "不是素数"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form2.Show
End Sub
