VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "IF�ۺ�����"
   ClientHeight    =   4140
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6645
   LinkTopic       =   "Form7"
   ScaleHeight     =   4140
   ScaleWidth      =   6645
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command4 
      Caption         =   "4���������������Ӵ�С�������"
      Height          =   735
      Left            =   3360
      TabIndex        =   3
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "3������һ������������С�����ж���ż��"
      Height          =   735
      Left            =   600
      TabIndex        =   2
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2�������������˷�  ����IF��"
      Height          =   735
      Left            =   3360
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1���ж����ݸİ�"
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim a!, b!
    a = Val(InputBox("������ߣ�cm��", "����", 170))
    b = Val(InputBox("�������أ�kg��", "����", 80))
    If b / (a * a) < 30 Then
        MsgBox "�㲻ƫ��"
    Else
        MsgBox "��ƫ���ˣ�"
    End If
End Sub

Private Sub Command2_Click()
    Dim a!
    a = Val(InputBox("����������", "����", 60))
    If a <= 50 Then
        MsgBox "�˷���" & a * 0.13
    Else
        MsgBox "�˷���" & 50 * 0.13 + (a - 50) * 0.2
    End If
End Sub

Private Sub Command3_Click()
    Dim a%
    a = InputBox("������һ������", "�ж���ż����", 100)
    If a Mod 2 = 0 Then
        MsgBox "��ż��"
    Else
        MsgBox "������"
    End If
End Sub

Private Sub Command4_Click()
    Dim a!, b!, c!, t!
    a = Val(InputBox("first number", "enter", 30))
    b = Val(InputBox("second number", "enter", 10.5))
    c = Val(InputBox("third number", "enter", 50))
    If a > b Then
        t = a
        a = b
        b = t
    End If
    If a > c Then
        t = a
        a = c
        c = t
    End If
    If b > c Then
        t = b
        b = c
        c = t
    End If
    Print "��������С����Ϊ" & a; b; c
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
