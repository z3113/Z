VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Discount"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7575
   LinkTopic       =   "Form4"
   ScaleHeight     =   4065
   ScaleWidth      =   7575
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Caption         =   "Unload"
      Height          =   495
      Left            =   6240
      TabIndex        =   0
      Top             =   3240
      Width           =   1215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a, b

Private Sub Command1_Click()
    Unload Form4
End Sub

Private Sub Form_Activate()
    a = Val(InputBox("�����빺��ĵ���", "����"))
    b = Val(InputBox("�����빺�������", "����"))
    If a * b <= 1000 Then
        Print "�㹺���˵���Ϊ" & a & "Ԫ����Ʒ" & b & "�������踶�ܼ�Ϊ"; a * b & "Ԫ���Żݺ�ʵ���踶" & a * b & "Ԫ"
    ElseIf a * b <= 2000 Then
        Print "�㹺���˵���Ϊ" & a & "Ԫ����Ʒ" & b & "�������踶�ܼ�Ϊ"; a * b & "Ԫ���Żݺ�ʵ���踶" & a * b * 0.9 & "Ԫ"
    ElseIf a * b <= 3000 Then
        Print "�㹺���˵���Ϊ" & a & "Ԫ����Ʒ" & b & "�������踶�ܼ�Ϊ"; a * b & "Ԫ���Żݺ�ʵ���踶" & a * b * 0.8 & "Ԫ"
    ElseIf a * b > 3000 Then
        Print "�㹺���˵���Ϊ" & a & "Ԫ����Ʒ" & b & "�������踶�ܼ�Ϊ"; a * b & "Ԫ���Żݺ�ʵ���踶" & a * b * 0.7 & "Ԫ"
    End If
End Sub
