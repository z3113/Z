VERSION 5.00
Begin VB.Form Form3 
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8655
   LinkTopic       =   "Form3"
   ScaleHeight     =   3630
   ScaleWidth      =   8655
   StartUpPosition =   3  '����ȱʡ
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    Dim a!, b!, c!, d!
    a = InputBox("�������봰��Ŀ�ȣ�" & Chr(10) & Chr(13) & "����ǣ�", "�������봰��Ŀ��", 1000)
    b = InputBox("�������봰��ĸ߶ȣ�" & vbCrLf & "�߶��ǣ�", "�������봰��ĸ߶�", 500)
    c = InputBox("�������봰�����߾࣡" & vbCrLf & "��߾��ǣ�", "�������봰�����߾�", 0)
    d = InputBox("�������봰����ұ߾࣡" & Chr(10) & Chr(13) & "�ұ߾��ǣ�", "�������봰����ұ߾�", 0)
    Form3.Width = Val(a)
    Form3.Height = Val(b)
    Form3.Left = Val(c)
    Form3.Top = Val(d)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Form3
    Form1.Show
End Sub
