VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "1��ż���ж�"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  '����ȱʡ
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim a%
    a = InputBox("������һ������", "�ж���ż����", 100)
    If a Mod 2 = 0 Then
        MsgBox "�������" & a & "��ż��"
    Else
        MsgBox "�������" & a & "������"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
