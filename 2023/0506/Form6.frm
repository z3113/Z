VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "����2���ɼ�����"
   ClientHeight    =   5430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8055
   LinkTopic       =   "Form6"
   ScaleHeight     =   5430
   ScaleWidth      =   8055
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command4 
      Caption         =   "����"
      Height          =   495
      Left            =   6480
      TabIndex        =   3
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ͳ�ƽ��"
      Height          =   495
      Left            =   6480
      TabIndex        =   2
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��ʾ�ɼ�"
      Height          =   495
      Left            =   6480
      TabIndex        =   1
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����ɼ�"
      Height          =   495
      Left            =   6480
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim a(10) As Integer, i As Integer
Private Sub Command1_Click()
    For i = 1 To 10
        a(i) = Val(InputBox("�������" & i & "��ͬѧ�ĳɼ�", "�ɼ�����", 60))
    Next i
End Sub

Private Sub Command2_Click()
    Print "10��ͬѧ�ĳɼ�Ϊ(���һ��)��"
    For i = 1 To 10
        Print a(i);
        If i Mod 5 = 0 Then Print
    Next i
End Sub

Private Sub Command3_Click()
    Dim b As Double, c As Integer
    For i = 1 To 10
        b = b + a(i)
    Next i
    b = b / 50
    For i = 1 To 10
        If a(i) > b Then c = c + 1
    Next i
    Print "ƽ����Ϊ" & b
    Print "����ƽ���ֵ�����" & c & "��"
End Sub

Private Sub Command4_Click()
    Unload Form6
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub
