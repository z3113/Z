VERSION 5.00
Begin VB.Form Form9 
   Caption         =   "���"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5475
   LinkTopic       =   "Form9"
   ScaleHeight     =   3060
   ScaleWidth      =   5475
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "����"
      Height          =   495
      Left            =   3480
      TabIndex        =   1
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "�Ӽ������һ��������5λ����������������Ǽ�λ����������λ���������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   1320
      TabIndex        =   0
      Top             =   1320
      Width           =   3735
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Form9
End Sub

Private Sub Form_Activate()
    Dim a%
    a = Val(InputBox("������һ��������5λ����������", "�ж���λ"))
    If a >= 10000 Then
        Print "������5λ��"
        Print StrReverse(a)
    ElseIf a >= 1000 Then
        Print "������4λ��"
        Print StrReverse(a)
    ElseIf a >= 100 Then
        Print "������3λ��"
        Print StrReverse(a)
    ElseIf a >= 10 Then
        Print "����λ2λ��"
        Print StrReverse(a)
    ElseIf a >= 0 Then
        Print "����Ϊ1λ��"
        Print StrReverse(a)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form2.Show
End Sub
